/**
 * Gujarat Form N — Leave Book (See rule 17) PDF layout.
 * Excel has three stacked sections; the generic grid mashed them into one table.
 */

export const FORM_NGJ_GJ_PDF_TITLE = 'FORM - N';
export const FORM_NGJ_GJ_PDF_RULE = '(See rule 17)';
export const FORM_NGJ_GJ_PDF_SUBTITLE = 'LEAVE BOOK';
export const FORM_NGJ_GJ_PDF_FESTIVAL_TITLE = 'DETAILS OF FESTIVAL LEAVE';
export const FORM_NGJ_GJ_PDF_CASUAL_TITLE = 'DETAILS OF CASUAL LEAVE';
export const FORM_NGJ_GJ_PDF_FOOTER = 'Name and Signature of Authority';

export const FORM_NGJ_GJ_PDF_LEAVE_LEAFS = [
  { key: 'leaveDueOn', label: 'Leave due on', parent: 'Accumulation of leave', number: '1' },
  { key: 'noOfDays', label: 'No. of days', parent: 'Accumulation of leave', number: '2' },
  { key: 'leaveFromTo', label: 'From - To -', parent: 'Leave allowed', number: '3' },
  { key: 'moiety1', label: '1st Moiety', parent: 'Payment for leave made on', number: '4' },
  { key: 'moiety2', label: '2nd Moiety', parent: 'Payment for leave made on', number: '4' },
  { key: 'applicationDate', label: 'Application Date', parent: 'Refusal of leave', number: '5' },
  { key: 'dateOfRefusal', label: 'Date of Refusal', parent: 'Refusal of leave', number: '5' },
  {
    key: 'dateOfDischarge',
    label: 'Date of discharge',
    parent: 'Payment for Leave on discharge of an worker quitting employment if admissible',
    number: '6',
  },
  {
    key: 'dateAndAmountPaid',
    label: 'Date and amount paid',
    parent: 'Payment for Leave on discharge of an worker quitting employment if admissible',
    number: '6',
  },
  {
    key: 'signature',
    label: 'Signature or thumb impression of worker',
    parent: 'Payment for Leave on discharge of an worker quitting employment if admissible',
    number: '6',
  },
  { key: 'leaveRemarks', label: 'Remarks', parent: 'Remarks', number: '7' },
];

export const FORM_NGJ_GJ_PDF_FESTIVAL_LEAFS = [
  { key: 'festivalPeriodFrom', label: 'From', parent: 'Period' },
  { key: 'festivalPeriodTo', label: 'To', parent: 'Period' },
  { key: 'festivalTotal', label: 'Total Leave', parent: 'Total Leave' },
  { key: 'festivalAvailed', label: 'Availed Leave', parent: 'Availed Leave' },
  { key: 'festivalBalance', label: 'Balance Leave', parent: 'Balance Leave' },
  {
    key: 'festivalPayment',
    label: 'Payment made in lieu of Festival Leave, when called for work',
    parent: 'Payment made in lieu of Festival Leave, when called for work',
  },
  { key: 'festivalRemarks', label: 'Remarks', parent: 'Remarks' },
];

export const FORM_NGJ_GJ_PDF_CASUAL_LEAFS = [
  { key: 'casualPeriodFrom', label: 'From', parent: 'Period' },
  { key: 'casualPeriodTo', label: 'To', parent: 'Period' },
  { key: 'casualTotal', label: 'Total Leave', parent: 'Total Leave' },
  { key: 'casualAvailed', label: 'Availed Leave', parent: 'Availed Leave' },
  { key: 'casualBalance', label: 'Balance Leave', parent: 'Balance Leave' },
  { key: 'casualRemarks', label: 'Remarks', parent: 'Remarks' },
];

const norm = (text) =>
  String(text || '')
    .replace(/\r\n/g, '\n')
    .replace(/\u00a0/g, ' ')
    .replace(/[ \t]+/g, ' ')
    .trim();

const lower = (text) => norm(text).toLowerCase();

export function looksLikeFormNGJGujaratPdfContext(
  metaLines,
  rows,
  sheetName = '',
  fileName = ''
) {
  const blob = [...(metaLines || []), ...(rows || []).slice(0, 28).flat(), sheetName || '', fileName || '']
    .join(' ')
    .toLowerCase();
  if (!blob) return false;
  if (/\bform\s*m\b/.test(blob) && /identity\s+card/.test(blob)) return false;
  if (
    (/\bform[\s._-]*l\b/.test(blob) || /form_l_gj/.test(blob)) &&
    /workers engaged in shift|see\s+rule\s+14/.test(blob)
  ) {
    return false;
  }
  const hasGujarat = /gujarat|\b_gj\b|form[\s._-]*n[\s._-]*gj|form_n_gj/.test(blob);
  const hasFormN =
    /\bform[\s._-]*n\b/.test(blob) || /form[\s._-]*n[\s._-]*gj/.test(blob) || /form_n_gj/.test(blob);
  if (hasGujarat && hasFormN) return true;
  if (hasFormN && /leave\s+book/.test(blob)) return true;
  if (hasFormN && /see\s+rule\s+17/.test(blob)) return true;
  return (
    /leave\s+book/.test(blob) &&
    /accumulation\s+of\s+leave/.test(blob) &&
    (/festival\s+leave/.test(blob) || /leave\s+due\s+on/.test(blob))
  );
}

function findRowIndex(rows, matcher) {
  for (let r = 0; r < (rows || []).length; r += 1) {
    const blob = (rows[r] || []).map((c) => lower(c)).join(' ');
    if (matcher(blob, rows[r] || [], r)) return r;
  }
  return -1;
}

function cellText(rows, r, c) {
  return norm(rows?.[r]?.[c]);
}

function parseLabeledBlock(text) {
  const raw = String(text || '');
  const out = {
    establishment: '',
    workerName: '',
    department: '',
    employer: '',
    dateOfEntry: '',
    receipt: '',
  };
  const specs = [
    { key: 'establishment', re: /name\s+of\s+the\s+establishment\s*:?\s*/i },
    { key: 'workerName', re: /name\s+of\s+the\s+worker\s*:?\s*/i },
    { key: 'department', re: /description\s+of\s+the\s+department(?:\s*\(if\s+applicable\))?\s*:?\s*/i },
    { key: 'employer', re: /name\s+of\s+the\s+employer\s*:?\s*/i },
    { key: 'dateOfEntry', re: /date\s+of\s+entry\s+into\s+service(?:\s*(?:worker\)|\([^)\n]*\)))?\s*:?\s*/i },
    { key: 'receipt', re: /receipt\s+of\s+le[av]e?\s*book\s*:?\s*/i },
  ];
  const lines = raw
    .split(/\n/)
    .map((l) => l.trim())
    .filter(Boolean);
  let currentKey = '';
  lines.forEach((line) => {
    if (/signature or thumb impression of worker/i.test(line) && !/date of entry/i.test(line)) {
      currentKey = '';
      return;
    }
    const spec = specs.find((s) => s.re.test(line));
    if (spec) {
      currentKey = spec.key;
      let value = line.replace(spec.re, '').replace(/^\(?if\s+applicable\)?:?\s*/i, '').trim();
      if (value) {
        out[spec.key] = out[spec.key] ? `${out[spec.key]}\n${value}` : value;
      }
      return;
    }
    if (!currentKey) return;
    let extra = line.replace(/^\(?if\s+applicable\)?:?\s*/i, '').trim();
    if (!extra) return;
    out[currentKey] = out[currentKey] ? `${out[currentKey]}\n${extra}` : extra;
  });
  return out;
}

function collectPreTableText(rows, endRow, colCount) {
  const parts = [];
  const last = endRow < 0 ? Math.min(12, (rows || []).length) : endRow;
  for (let r = 0; r < last; r += 1) {
    for (let c = 0; c < Math.max(1, colCount || 12); c += 1) {
      const t = cellText(rows, r, c);
      if (t) parts.push(t);
    }
  }
  return parts.join('\n');
}

function findLeafColumns(rows, startRow, endRow, colCount, leafSpecs) {
  const cols = leafSpecs.map((spec) => ({ ...spec, col: -1 }));
  const maxR = Math.min((rows || []).length - 1, endRow);
  for (let r = Math.max(0, startRow); r <= maxR; r += 1) {
    for (let c = 0; c < colCount; c += 1) {
      const t = lower(cellText(rows, r, c));
      if (!t) continue;
      cols.forEach((spec) => {
        if (spec.col >= 0) return;
        const re = new RegExp(spec.match || spec.label.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'i');
        if (re.test(t)) spec.col = c;
      });
    }
  }
  return cols;
}

function isFormNGJHeaderLikeRow(blob) {
  const t = lower(blob);
  if (!t) return true;
  return (
    /form\s*-?\s*n\b|see\s+rule\s+17|leave\s+book/.test(t) ||
    /accumulation\s+of\s+leave|leave\s+allowed|payment\s+for\s+leave|refusal\s+of\s+leave/.test(t) ||
    /details\s+of\s+(festival|casual)\s+leave/.test(t) ||
    /leave\s+due\s+on|no\.?\s+of\s+days|from\s*-+\s*to|1st\s+moiety|2nd\s+moiety/.test(t) ||
    /application\s+date|date\s+of\s+refusal|date\s+of\s+discharge|amount\s+paid|thumb\s+impression/.test(t) ||
    /^(from|to|period)(\s+(from|to|period))*$/.test(t) ||
    /total\s+leave|availed\s+leave|balance\s+leave|payment\s+made\s+in\s+lieu/.test(t) ||
    /name\s+and\s+signature\s+of\s+authority/.test(t) ||
    /name\s+of\s+the\s+(establishment|worker|employer)|description\s+of\s+the\s+department|date\s+of\s+entry/.test(
      t
    )
  );
}

function firstDataRow(rows, afterRow, endRow, colCount) {
  const stop = Math.min((rows || []).length - 1, endRow);
  for (let r = Math.max(0, afterRow); r <= stop; r += 1) {
    const values = [];
    for (let c = 0; c < colCount; c += 1) values.push(cellText(rows, r, c));
    const blob = values.join(' ');
    if (!blob.trim()) continue;
    if (isFormNGJHeaderLikeRow(blob)) continue;
    if (values.every((c) => !c || /^\d+$/.test(c))) continue;
    return { rowIndex: r, values };
  }
  return { rowIndex: -1, values: [] };
}

function valuesForLeafs(leafCols, dataValues) {
  const out = {};
  leafCols.forEach((spec) => {
    out[spec.key] = spec.col >= 0 ? String(dataValues[spec.col] || '').trim() : '';
  });
  return out;
}

function emptyValues(leafs) {
  const out = {};
  leafs.forEach((spec) => {
    out[spec.key] = '';
  });
  return out;
}

export function normalizeFormNGJGujaratPdfMatrix(rows, colCount, tableStartRow = 0, metaLines = []) {
  const src = Array.isArray(rows) ? rows : [];
  const cols = Math.max(1, Number(colCount) || 12);
  const leaveRow = findRowIndex(src, (blob) => /accumulation\s+of\s+leave/.test(blob));
  const festivalRow = findRowIndex(src, (blob) => /details\s+of\s+festival\s+leave/.test(blob));
  const casualRow = findRowIndex(src, (blob) => /details\s+of\s+casual\s+leave/.test(blob));

  const parsed = parseLabeledBlock(collectPreTableText(src, leaveRow < 0 ? 12 : leaveRow, cols));
  const metaParsed = parseLabeledBlock((metaLines || []).join('\n'));
  const identity = {
    establishment: parsed.establishment || metaParsed.establishment,
    workerName: parsed.workerName || metaParsed.workerName,
    department: parsed.department || metaParsed.department,
    employer: parsed.employer || metaParsed.employer,
    dateOfEntry: parsed.dateOfEntry || metaParsed.dateOfEntry,
    receipt: parsed.receipt || metaParsed.receipt,
  };

  const leaveEnd = festivalRow >= 0 ? festivalRow - 1 : casualRow >= 0 ? casualRow - 1 : src.length - 1;
  const festivalEnd = casualRow >= 0 ? casualRow - 1 : src.length - 1;

  const leaveLeafs = FORM_NGJ_GJ_PDF_LEAVE_LEAFS.map((spec) => ({
    ...spec,
    match: spec.label.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'),
  }));
  leaveLeafs[2].match = 'from\\s*-?\\s*to';
  leaveLeafs[3].match = '1st\\s+moiety';
  leaveLeafs[4].match = '2nd\\s+moiety';
  const leaveCols = findLeafColumns(
    src,
    leaveRow < 0 ? tableStartRow : leaveRow,
    leaveEnd,
    cols,
    leaveLeafs
  );
  const leaveData = firstDataRow(src, leaveRow < 0 ? tableStartRow : leaveRow, leaveEnd, cols);

  const festivalCols = findLeafColumns(src, festivalRow, festivalEnd, cols, [
    { key: 'festivalPeriodFrom', match: '^from$' },
    { key: 'festivalPeriodTo', match: '^to$' },
    { key: 'festivalTotal', match: 'total\\s+leave' },
    { key: 'festivalAvailed', match: 'availed\\s+leave' },
    { key: 'festivalBalance', match: 'balance\\s+leave' },
    { key: 'festivalPayment', match: 'payment\\s+made\\s+in\\s+lieu' },
    { key: 'festivalRemarks', match: '^remarks$' },
  ]);
  const festivalData = firstDataRow(src, festivalRow < 0 ? leaveEnd + 1 : festivalRow, festivalEnd, cols);

  const casualCols = findLeafColumns(src, casualRow, src.length - 1, cols, [
    { key: 'casualPeriodFrom', match: '^from$' },
    { key: 'casualPeriodTo', match: '^to$' },
    { key: 'casualTotal', match: 'total\\s+leave' },
    { key: 'casualAvailed', match: 'availed\\s+leave' },
    { key: 'casualBalance', match: 'balance\\s+leave' },
    { key: 'casualRemarks', match: '^remarks$' },
  ]);
  const casualData = firstDataRow(
    src,
    casualRow < 0 ? festivalEnd + 1 : casualRow,
    src.length - 1,
    cols
  );

  const model = {
    titles: [FORM_NGJ_GJ_PDF_TITLE, FORM_NGJ_GJ_PDF_RULE, FORM_NGJ_GJ_PDF_SUBTITLE],
    identity,
    leaveValues: leaveData.values.length
      ? valuesForLeafs(leaveCols, leaveData.values)
      : emptyValues(FORM_NGJ_GJ_PDF_LEAVE_LEAFS),
    festivalValues: festivalData.values.length
      ? valuesForLeafs(festivalCols, festivalData.values)
      : emptyValues(FORM_NGJ_GJ_PDF_FESTIVAL_LEAFS),
    casualValues: casualData.values.length
      ? valuesForLeafs(casualCols, casualData.values)
      : emptyValues(FORM_NGJ_GJ_PDF_CASUAL_LEAFS),
    footer: FORM_NGJ_GJ_PDF_FOOTER,
  };

  return {
    rows: src,
    colCount: cols,
    tableStartRow: leaveRow >= 0 ? leaveRow : tableStartRow,
    metaLines: model.titles,
    formNGJLayout: true,
    formNGJModel: model,
  };
}

export function applyFormNGJGujaratPdfNormalization(target, normalized) {
  if (!target || !normalized) return target;
  target.rows = normalized.rows;
  target.colCount = normalized.colCount;
  target.tableStartRow = normalized.tableStartRow;
  target.metaLines = Array.isArray(normalized.metaLines) ? normalized.metaLines : [];
  target.formNGJLayout = true;
  target.formNGJModel = normalized.formNGJModel;
  return target;
}

function strokeRect(doc, x, y, w, h) {
  doc.setDrawColor(0, 0, 0);
  doc.setLineWidth(0.6);
  doc.rect(x, y, w, h, 'S');
}

function drawWrapped(doc, text, x, y, w, size, align, bold) {
  doc.setFont('helvetica', bold ? 'bold' : 'normal');
  doc.setFontSize(size);
  const lines = doc.splitTextToSize(String(text || ''), Math.max(12, w - 4));
  const lineH = size + 2;
  const top = y + size;
  if (align === 'center') {
    doc.text(lines, x + w / 2, top, { align: 'center' });
  } else {
    doc.text(lines, x + 2, top);
  }
  return lines.length * lineH;
}

function ensurePage(doc, y, need, marginBottom = 36) {
  const pageHeight = doc.internal.pageSize.getHeight();
  if (y + need <= pageHeight - marginBottom) return y;
  doc.addPage();
  return 28;
}

function colWidths(total, weights) {
  const sum = weights.reduce((a, b) => a + b, 0) || 1;
  return weights.map((w) => (total * w) / sum);
}

function drawHeaderDataTable(doc, x, y, width, groups, values) {
  const leafs = [];
  groups.forEach((g) =>
    g.leafs.forEach((leaf) =>
      leafs.push({ ...leaf, parent: g.label, number: leaf.number || g.number || '' })
    )
  );
  const widths = colWidths(
    width,
    leafs.map((leaf) => leaf.weight || 8)
  );
  const parentH = 42;
  const numH = 18;
  const leafH = 36;
  const dataH = 28;
  let cursor = x;
  groups.forEach((group) => {
    const spanW = group.leafs.reduce((sum, leaf) => {
      const leafIndex = leafs.findIndex((l) => l.key === leaf.key);
      return sum + (widths[leafIndex] || 0);
    }, 0);
    strokeRect(doc, cursor, y, spanW, parentH);
    drawWrapped(doc, group.label, cursor, y + 3, spanW, 9, 'center', true);
    cursor += spanW;
  });
  cursor = x;
  let i = 0;
  while (i < leafs.length) {
    let span = 1;
    while (i + span < leafs.length && leafs[i + span].number === leafs[i].number) span += 1;
    const spanW = widths.slice(i, i + span).reduce((sum, w) => sum + w, 0);
    strokeRect(doc, cursor, y + parentH, spanW, numH);
    drawWrapped(doc, leafs[i].number || '', cursor, y + parentH + 1, spanW, 11, 'center', true);
    cursor += spanW;
    i += span;
  }
  cursor = x;
  leafs.forEach((leaf, idx) => {
    strokeRect(doc, cursor, y + parentH + numH, widths[idx], leafH);
    drawWrapped(doc, leaf.label, cursor, y + parentH + numH + 3, widths[idx], 9, 'center', true);
    cursor += widths[idx];
  });
  const dataY = y + parentH + numH + leafH;
  cursor = x;
  leafs.forEach((leaf, idx) => {
    strokeRect(doc, cursor, dataY, widths[idx], dataH);
    drawWrapped(doc, values?.[leaf.key] || '', cursor, dataY + 5, widths[idx], 10, 'center', false);
    cursor += widths[idx];
  });
  return dataY + dataH;
}

function drawSimpleSectionTable(doc, x, y, width, title, groups, values) {
  doc.setFont('helvetica', 'bold');
  doc.setFontSize(12);
  doc.text(title, x, y + 13);
  let nextY = y + 18;
  const leafs = [];
  groups.forEach((g) => g.leafs.forEach((leaf) => leafs.push({ ...leaf, parent: g.label })));
  const widths = colWidths(
    width,
    leafs.map((leaf) => leaf.weight || 10)
  );
  const parentH = 38;
  const leafH = 20;
  const dataH = 26;
  let cursor = x;
  groups.forEach((group) => {
    const spanW = group.leafs.reduce((sum, leaf) => {
      const leafIndex = leafs.findIndex((l) => l.key === leaf.key);
      return sum + (widths[leafIndex] || 0);
    }, 0);
    strokeRect(doc, cursor, nextY, spanW, parentH);
    drawWrapped(doc, group.label, cursor, nextY + 3, spanW, 9.5, 'center', true);
    cursor += spanW;
  });
  cursor = x;
  leafs.forEach((leaf, i) => {
    if (leaf.label && leaf.label !== leaf.parent) {
      strokeRect(doc, cursor, nextY + parentH, widths[i], leafH);
      drawWrapped(doc, leaf.label, cursor, nextY + parentH + 2, widths[i], 9, 'center', true);
    } else {
      strokeRect(doc, cursor, nextY + parentH, widths[i], leafH);
    }
    cursor += widths[i];
  });
  const dataY = nextY + parentH + leafH;
  cursor = x;
  leafs.forEach((leaf, i) => {
    strokeRect(doc, cursor, dataY, widths[i], dataH);
    drawWrapped(doc, values?.[leaf.key] || '', cursor, dataY + 5, widths[i], 10, 'center', false);
    cursor += widths[i];
  });
  return dataY + dataH;
}

/**
 * Draw Form N as the Excel Leave Book: identity boxes + 3 stacked tables.
 */
export function drawFormNGJGujaratLeaveBook(doc, matrix, startY = 24) {
  const pageWidth = doc.internal.pageSize.getWidth();
  const marginX = 28;
  const usableWidth = pageWidth - marginX * 2;
  const model = matrix?.formNGJModel || normalizeFormNGJGujaratPdfMatrix(matrix?.rows || [], matrix?.colCount || 12)
    .formNGJModel;
  let y = startY;

  (model.titles || []).forEach((title) => {
    const isName = /^form\s*[-–]?\s*n\b/i.test(title);
    const isBook = /leave\s+book/i.test(title);
    doc.setFont('helvetica', isName || isBook ? 'bold' : 'normal');
    doc.setFontSize(isName ? 16 : isBook ? 13 : 11);
    doc.text(String(title || ''), pageWidth / 2, y + (isName ? 16 : 12), { align: 'center' });
    y += isName ? 20 : 15;
  });
  y += 8;

  const id = model.identity || {};
  const leftLines = [
    `Name of the establishment: ${id.establishment || ''}`.trim(),
    `Name of the worker: ${id.workerName || ''}`.trim(),
    `Description of the Department (if applicable): ${id.department || ''}`.trim(),
  ];
  const rightLines = [
    `Name of the employer: ${id.employer || ''}`.trim(),
    `Date of entry into service: ${id.dateOfEntry || ''}`.trim(),
    '(Signature or thumb impression of worker)',
  ];
  const half = usableWidth / 2;
  doc.setFont('helvetica', 'normal');
  doc.setFontSize(11);
  const wrapBlock = (lines) =>
    lines.flatMap((line) => doc.splitTextToSize(String(line || ''), half - 10));
  const leftWrapped = wrapBlock(leftLines);
  const rightWrapped = wrapBlock(rightLines);
  const boxH = Math.max(78, Math.max(leftWrapped.length, rightWrapped.length) * 14 + 14);
  strokeRect(doc, marginX, y, half, boxH);
  strokeRect(doc, marginX + half, y, half, boxH);
  doc.text(leftWrapped, marginX + 5, y + 16);
  doc.text(rightWrapped, marginX + half + 5, y + 16);
  y += boxH + 14;

  const leaveGroups = [
    {
      label: 'Accumulation of leave',
      number: '',
      leafs: [
        { key: 'leaveDueOn', label: 'Leave due on', number: '1', weight: 8 },
        { key: 'noOfDays', label: 'No. of days', number: '2', weight: 7 },
      ],
    },
    {
      label: 'Leave allowed',
      leafs: [{ key: 'leaveFromTo', label: 'From - To -', number: '3', weight: 10 }],
    },
    {
      label: 'Payment for leave made on',
      leafs: [
        { key: 'moiety1', label: '1st Moiety', number: '4', weight: 7 },
        { key: 'moiety2', label: '2nd Moiety', number: '4', weight: 7 },
      ],
    },
    {
      label: 'Refusal of leave',
      leafs: [
        { key: 'applicationDate', label: 'Application Date', number: '5', weight: 9 },
        { key: 'dateOfRefusal', label: 'Date of Refusal', number: '5', weight: 8 },
      ],
    },
    {
      label: 'Payment for Leave on discharge of an worker quitting employment if admissible',
      leafs: [
        { key: 'dateOfDischarge', label: 'Date of discharge', number: '6', weight: 8 },
        { key: 'dateAndAmountPaid', label: 'Date and amount paid', number: '6', weight: 9 },
        { key: 'signature', label: 'Signature or thumb impression of worker', number: '6', weight: 11 },
      ],
    },
    {
      label: 'Remarks',
      leafs: [{ key: 'leaveRemarks', label: 'Remarks', number: '7', weight: 8 }],
    },
  ];
  y = ensurePage(doc, y, 130);
  y = drawHeaderDataTable(doc, marginX, y, usableWidth, leaveGroups, model.leaveValues);
  y += 20;

  const festivalWidth = usableWidth * 0.78;
  y = ensurePage(doc, y, 110);
  y = drawSimpleSectionTable(
    doc,
    marginX,
    y,
    festivalWidth,
    FORM_NGJ_GJ_PDF_FESTIVAL_TITLE,
    [
      {
        label: 'Period',
        leafs: [
          { key: 'festivalPeriodFrom', label: 'From', weight: 10 },
          { key: 'festivalPeriodTo', label: 'To', weight: 10 },
        ],
      },
      { label: 'Total Leave', leafs: [{ key: 'festivalTotal', label: 'Total Leave', weight: 9 }] },
      { label: 'Availed Leave', leafs: [{ key: 'festivalAvailed', label: 'Availed Leave', weight: 9 }] },
      { label: 'Balance Leave', leafs: [{ key: 'festivalBalance', label: 'Balance Leave', weight: 9 }] },
      {
        label: 'Payment made in lieu of Festival Leave, when called for work',
        leafs: [{ key: 'festivalPayment', label: 'Payment made in lieu of Festival Leave, when called for work', weight: 16 }],
      },
      { label: 'Remarks', leafs: [{ key: 'festivalRemarks', label: 'Remarks', weight: 9 }] },
    ],
    model.festivalValues
  );
  y += 20;

  const casualWidth = usableWidth * 0.62;
  y = ensurePage(doc, y, 110);
  y = drawSimpleSectionTable(
    doc,
    marginX,
    y,
    casualWidth,
    FORM_NGJ_GJ_PDF_CASUAL_TITLE,
    [
      {
        label: 'Period',
        leafs: [
          { key: 'casualPeriodFrom', label: 'From', weight: 10 },
          { key: 'casualPeriodTo', label: 'To', weight: 10 },
        ],
      },
      { label: 'Total Leave', leafs: [{ key: 'casualTotal', label: 'Total Leave', weight: 9 }] },
      { label: 'Availed Leave', leafs: [{ key: 'casualAvailed', label: 'Availed Leave', weight: 9 }] },
      { label: 'Balance Leave', leafs: [{ key: 'casualBalance', label: 'Balance Leave', weight: 9 }] },
      { label: 'Remarks', leafs: [{ key: 'casualRemarks', label: 'Remarks', weight: 9 }] },
    ],
    model.casualValues
  );

  y += 26;
  y = ensurePage(doc, y, 24);
  doc.setFont('helvetica', 'normal');
  doc.setFontSize(11);
  doc.text(FORM_NGJ_GJ_PDF_FOOTER, marginX + usableWidth, y, { align: 'right' });
  return y + 8;
}
