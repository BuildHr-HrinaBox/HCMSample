/**
 * Karnataka Form F — Register of Leave with Wages (See Rule 8).
 * Excel is identity fields + PART I earned-leave grid + PART II sick/accident.
 * The generic PDF grid treated those bands as one 11-col table.
 */

export const FORM_F_KA_PDF_TITLE = 'FORM F';
export const FORM_F_KA_PDF_RULE = '(SEE RULE 8)';
export const FORM_F_KA_PDF_SUBTITLE = 'REGISTER OF LEAVE WITH WAGES';
export const FORM_F_KA_PDF_PART_I_TITLE = 'PART I EARNED LEAVE';
export const FORM_F_KA_PDF_PART_II_TITLE = 'PART II - Sick/Accident Leave (with pay)';

export const FORM_F_KA_PDF_IDENTITY_FIELDS = [
  {
    key: 'slNo',
    number: '1',
    label: 'SI No in the Register of Adult/young person',
    match: (n) =>
      (/s[li]\s*no/.test(n) && /register/.test(n)) ||
      /s[li]\s*no\s+in\s+the/.test(n),
  },
  {
    key: 'dateOfEntry',
    number: '2',
    label: 'Date of entry into service',
    match: (n) => /date(?:\s+of)?\s+entry\s+into\s+service/.test(n) || /entry\s+into\s+service/.test(n),
  },
  {
    key: 'personName',
    number: '3',
    label: 'Name of the person',
    match: (n) => {
      if (/father/.test(n) || /husband/.test(n) || /establishment/.test(n)) return false;
      return /name\s+of\s+the\s+person/.test(n) || /^\d+\.?\s*name\s+of\s+the/.test(n);
    },
  },
  {
    key: 'fatherName',
    number: '4',
    label: "Father's Name",
    match: (n) => (/father/.test(n) && /name/.test(n)) || /husband'?s\s+name/.test(n),
  },
];

export const FORM_F_KA_PDF_PART_I_GROUPS = [
  {
    label: 'No of Days worked',
    leafs: [
      { key: 'from', label: 'From', number: '1', weight: 10 },
      { key: 'to', label: 'To', number: '2', weight: 10 },
      { key: 'totalDays', label: 'Total days worked', number: '3', weight: 8 },
    ],
  },
  {
    label: 'Leave earned',
    leafs: [{ key: 'leaveEarned', label: '', number: '4', weight: 7 }],
  },
  {
    label: 'Leave at credit (incl balance if any, on return from leave on last occasion)',
    leafs: [{ key: 'leaveAtCredit', label: '', number: '5', weight: 14 }],
  },
  {
    label: 'leave availed',
    leafs: [
      { key: 'availedFrom', label: 'From', number: '6', weight: 10 },
      { key: 'availedTo', label: 'To', number: '7', weight: 10 },
      { key: 'availedDays', label: 'No. of days', number: '8', weight: 7 },
    ],
  },
  {
    label: 'Balance on return from leave',
    leafs: [{ key: 'balance', label: '', number: '9', weight: 9 }],
  },
  {
    label: 'Date on which wages for leave paid and amount paid',
    leafs: [{ key: 'wagesPaid', label: '', number: '10', weight: 13 }],
  },
  {
    label: 'Remarks',
    leafs: [{ key: 'remarks', label: '', number: '11', weight: 8 }],
  },
];

export const FORM_F_KA_PDF_PART_I_LEAFS = FORM_F_KA_PDF_PART_I_GROUPS.flatMap((g) => g.leafs);

export const FORM_F_KA_PDF_PART_II_GROUPS = [
  {
    label: '',
    leafs: [{ key: 'year', label: 'Year', number: '1', weight: 10 }],
  },
  {
    label: 'Sick/Accident leave',
    leafs: [
      { key: 'credit', label: 'of credit', number: '2', weight: 12 },
      { key: 'availed', label: 'Availed', number: '3', weight: 12 },
    ],
  },
  {
    label: '',
    leafs: [{ key: 'balance', label: 'Balance', number: '4', weight: 10 }],
  },
];

const MIN_PART_I_BODY_ROWS = 8;
const MIN_PART_II_BODY_ROWS = 6;

const norm = (value) =>
  String(value ?? '')
    .replace(/_x000d_/gi, '\n')
    .replace(/\u000d/g, '\n')
    .replace(/\r\n/g, '\n')
    .replace(/\r/g, '\n')
    .replace(/\u00a0/g, ' ')
    .replace(/\[object Object\]/gi, '')
    .replace(/[ \t]+/g, ' ')
    .trim();

const lower = (value) => norm(value).toLowerCase();

const blobFrom = (metaLines, rows, sheetName = '', fileName = '') =>
  [...(metaLines || []), ...(rows || []).slice(0, 28).flat(), sheetName || '', fileName || '']
    .map((t) => lower(t))
    .join(' ');

const isFormFFilename = (sheetName = '', fileName = '') => {
  const name = `${sheetName || ''} ${fileName || ''}`.toLowerCase();
  if (/form[\s._-]*15\b/.test(name) || /form[\s._-]*xv\b/.test(name)) return false;
  return (
    /form[\s._-]*f[\s._-]*ka\b/.test(name) ||
    /form[\s._-]*f[\s._-]*karnataka/.test(name) ||
    /form_f_-_karnataka/.test(name) ||
    /form_f_karnataka/.test(name)
  );
};

const isForm15LeaveRegisterBlob = (blob) =>
  (/\bform[\s._-]*15\b/.test(blob) || /\bform[\s._-]*xv\b/.test(blob)) &&
  !/\bform[\s._-]*f\b/.test(blob);

export const looksLikeFormFKarnatakaPdfContext = (
  metaLines = [],
  rows = [],
  sheetName = '',
  fileName = ''
) => {
  if (isFormFFilename(sheetName, fileName)) return true;
  const blob = blobFrom(metaLines, rows, sheetName, fileName);
  if (!blob) return false;
  if (isForm15LeaveRegisterBlob(blob)) return false;
  if (/appointment\s+order/.test(blob) && /see\s+rule\s*24/.test(blob)) return false;
  if (/combined\s+muster\s+roll/.test(blob)) return false;

  const hasFormF = /(?:^|[^a-z0-9])form[\s._-]*f(?:[^a-z0-9]|$)/.test(blob);
  const hasLeaveRegister =
    /register\s+of\s+leave\s+with\s+wages/.test(blob) || /leave\s+with\s+wages/.test(blob);
  const hasRule8 = /see\s+rule\s*8\b/.test(blob);
  const hasPartI = /part\s*[-–]?\s*i\b/.test(blob) && /earned\s+leave/.test(blob);
  const hasPartIi = /part\s*[-–]?\s*ii/.test(blob) && /sick|accident/.test(blob);
  const hasKa = /karnataka/.test(blob) || /form[\s._-]*f[\s._-]*ka/.test(blob);

  if (hasFormF && (hasLeaveRegister || hasRule8 || hasPartI)) return true;
  if (hasFormF && hasKa) return true;
  if (hasLeaveRegister && hasRule8 && (hasPartI || hasPartIi)) return true;
  return false;
};

const cellText = (rows, r, c) => norm(rows?.[r]?.[c]);

const rowBlob = (row) =>
  (Array.isArray(row) ? row : [])
    .map((c) => lower(c))
    .filter(Boolean)
    .join(' ');

const isTitleText = (text) => {
  const t = lower(text);
  if (!t) return false;
  if (/^form\s*f$/.test(t) || /^form\s*f\b/.test(t) && t.length < 18) return true;
  if (/see\s+rule\s*8\b/.test(t)) return true;
  if (/register\s+of\s+leave\s+with\s+wages/.test(t)) return true;
  return false;
};

const isPartITitle = (text) => {
  const t = lower(text);
  return /part\s*[-–]?\s*i\b/.test(t) && /earned\s+leave/.test(t);
};

const isPartIiTitle = (text) => {
  const t = lower(text);
  return /part\s*[-–]?\s*ii/.test(t) && (/sick|accident/.test(t) || t.length < 40);
};

const isSystemGeneratedNote = (text) =>
  /system\s+generated\s+document|computer\s+generated\s+document/.test(lower(text));

const isColumnIndexRow = (row) => {
  const filled = (Array.isArray(row) ? row : []).map((c) => norm(c)).filter(Boolean);
  if (filled.length < 6) return false;
  const nums = filled.filter((t) => /^\d{1,2}$/.test(t));
  return nums.length >= 6 && nums.length / filled.length >= 0.7;
};

const isPartIHeaderRow = (row) => {
  const blob = rowBlob(row);
  if (!blob) return false;
  if (isPartITitle(blob) || isPartIiTitle(blob) || isTitleText(blob)) return false;
  return (
    /no\s+of\s+days\s+worked/.test(blob) ||
    (/leave\s+earned/.test(blob) && /leave\s+availed|leave\s+at\s+credit/.test(blob)) ||
    (/^from$/.test(lower(row?.[0])) && /total\s+days\s+worked/.test(blob))
  );
};

const findRowIndex = (rows, matcher) => {
  for (let r = 0; r < (rows || []).length; r += 1) {
    if (matcher(rowBlob(rows[r]), rows[r] || [], r)) return r;
  }
  return -1;
};

const valueAfterColon = (text) => {
  const src = norm(text);
  const idx = src.indexOf(':');
  if (idx < 0) return '';
  return src.slice(idx + 1).replace(/^[:\s]+/, '').trim();
};

const isBareNumber = (text) => /^\d{1,2}$/.test(norm(text));

const looksLikeIdentityValue = (text, specKey) => {
  const t = norm(text);
  if (!t || isBareNumber(t) || isTitleText(t) || isPartITitle(t) || isPartIiTitle(t)) return false;
  if (FORM_F_KA_PDF_IDENTITY_FIELDS.some((s) => s.match(lower(t)))) return false;
  if (specKey === 'slNo') return /[a-z0-9]/i.test(t) && t.length <= 24;
  if (specKey === 'dateOfEntry') return /\d/.test(t);
  if (specKey === 'personName') return /[a-z]/i.test(t) && !/father/i.test(t);
  if (specKey === 'fatherName') return /[a-z]/i.test(t);
  return t.length > 0;
};

const extractIdentity = (rows, colCount) => {
  const values = { slNo: '', dateOfEntry: '', personName: '', fatherName: '' };
  const last = Math.min((rows || []).length, 22);
  for (let r = 0; r < last; r += 1) {
    const row = rows[r] || [];
    for (let c = 0; c < Math.max(colCount, row.length); c += 1) {
      const raw = cellText(rows, r, c);
      if (!raw) continue;
      const n = lower(raw);
      const spec = FORM_F_KA_PDF_IDENTITY_FIELDS.find((s) => s.match(n));
      if (!spec || values[spec.key]) continue;
      const inline = valueAfterColon(raw);
      if (inline && looksLikeIdentityValue(inline, spec.key)) {
        values[spec.key] = inline;
        continue;
      }
      for (let nc = c + 1; nc <= Math.min(c + 4, Math.max(colCount, row.length) - 1); nc += 1) {
        const beside = cellText(rows, r, nc);
        if (!beside) continue;
        if (FORM_F_KA_PDF_IDENTITY_FIELDS.some((s) => s.match(lower(beside)))) break;
        if (looksLikeIdentityValue(beside, spec.key)) {
          values[spec.key] = beside;
          break;
        }
      }
      if (!values[spec.key] && r + 1 < last) {
        const below = cellText(rows, r + 1, c);
        if (looksLikeIdentityValue(below, spec.key)) values[spec.key] = below;
      }
    }
  }
  return values;
};

const packPartIRow = (row, startCol) => {
  const src = Array.isArray(row) ? row : [];
  const slice = src.slice(startCol, startCol + FORM_F_KA_PDF_PART_I_LEAFS.length).map((c) => norm(c));
  const filled = slice.filter(Boolean);
  if (filled.length >= 3) return slice;
  const packed = src.map((c) => norm(c)).filter(Boolean);
  if (packed.length && packed.length <= FORM_F_KA_PDF_PART_I_LEAFS.length) {
    const out = Array(FORM_F_KA_PDF_PART_I_LEAFS.length).fill('');
    packed.forEach((v, i) => {
      out[i] = v;
    });
    return out;
  }
  return slice;
};

const isPartIDataRow = (row) => {
  const blob = rowBlob(row);
  if (!blob) return false;
  if (isTitleText(blob) || isPartITitle(blob) || isPartIiTitle(blob)) return false;
  if (isColumnIndexRow(row) || isPartIHeaderRow(row)) return false;
  if (isSystemGeneratedNote(blob)) return false;
  if (FORM_F_KA_PDF_IDENTITY_FIELDS.some((s) => s.match(blob))) return false;
  const filled = (Array.isArray(row) ? row : []).map((c) => norm(c)).filter(Boolean);
  return filled.some((t) => /\d/.test(t) || /[a-z]/i.test(t));
};

const extractPartIRows = (rows, colCount) => {
  const partIRow = findRowIndex(rows, (blob) => isPartITitle(blob));
  const partIiRow = findRowIndex(rows, (blob) => isPartIiTitle(blob));
  const scanFrom = partIRow >= 0 ? partIRow + 1 : 0;
  const scanTo = partIiRow >= 0 ? partIiRow : (rows || []).length;

  let numberRow = -1;
  let startCol = 0;
  for (let r = scanFrom; r < scanTo; r += 1) {
    const row = rows[r] || [];
    if (!isColumnIndexRow(row)) continue;
    const idx = row.findIndex((c) => /^\s*1\s*$/.test(norm(c)));
    if (idx >= 0) startCol = idx;
    numberRow = r;
    break;
  }

  const dataStart = numberRow >= 0 ? numberRow + 1 : scanFrom;
  const out = [];
  for (let r = dataStart; r < scanTo; r += 1) {
    const row = rows[r] || [];
    if (!isPartIDataRow(row)) continue;
    const packed = packPartIRow(row, startCol);
    if (packed.some((v) => v)) out.push(packed);
  }
  void colCount;
  return out;
};

const isPartIiLabel = (text) => {
  const t = lower(text);
  if (!t) return false;
  return (
    /^year$/.test(t) ||
    /of\s+credit/.test(t) ||
    /^credit$/.test(t) ||
    /^availed$/.test(t) ||
    /sick\/?accident/.test(t) ||
    /balance/.test(t) ||
    isPartIiTitle(t) ||
    isBareNumber(t)
  );
};

const extractPartIi = (rows, colCount) => {
  const values = { year: '', credit: '', availed: '', balance: '' };
  const partIiRow = findRowIndex(rows, (blob) => isPartIiTitle(blob));
  if (partIiRow < 0) return values;
  const scanTo = Math.min((rows || []).length, partIiRow + 10);
  const cells = [];
  for (let r = partIiRow + 1; r < scanTo; r += 1) {
    const row = rows[r] || [];
    for (let c = 0; c < Math.max(colCount, row.length); c += 1) {
      const t = cellText(rows, r, c);
      if (t) cells.push({ r, c, t, n: lower(t) });
    }
  }
  const yearHit = cells.find((x) => /^(19|20)\d{2}$/.test(x.t));
  if (yearHit) values.year = yearHit.t;

  const creditLabel = cells.find((x) => /credit/.test(x.n));
  if (creditLabel) {
    const beside = cells.find(
      (x) => x.r === creditLabel.r && x.c > creditLabel.c && x.c <= creditLabel.c + 2 && !isPartIiLabel(x.t)
    );
    if (beside) values.credit = beside.t;
  }
  const availedLabel = cells.find((x) => /^availed$/.test(x.n) || (/availed/.test(x.n) && !/leave/.test(x.n)));
  if (availedLabel) {
    const beside = cells.find(
      (x) => x.r === availedLabel.r && x.c > availedLabel.c && x.c <= availedLabel.c + 2 && !isPartIiLabel(x.t)
    );
    if (beside) values.availed = beside.t;
  }
  const balanceLabel = cells.find((x) => /balance/.test(x.n));
  if (balanceLabel) {
    const beside = cells.find(
      (x) => x.r === balanceLabel.r && x.c > balanceLabel.c && x.c <= balanceLabel.c + 2 && !isPartIiLabel(x.t)
    );
    if (beside) values.balance = beside.t;
  }

  if (!values.balance) {
    const nums = cells
      .filter((x) => /^-?\d+(\.\d+)?$/.test(x.t) && !/^(19|20)\d{2}$/.test(x.t) && !isBareNumber(x.t))
      .concat(cells.filter((x) => /^-?\d+(\.\d+)?$/.test(x.t) && !/^(19|20)\d{2}$/.test(x.t)));
    const leftover = nums.find((x) => !isBareNumber(x.t) || Number(x.t) === 0 || Number(x.t) > 11);
    if (leftover && leftover.t !== values.year) values.balance = leftover.t;
  }

  // Header-row model from the Excel screenshot: year | of credit | Availed | balance.
  if (!values.year || values.balance === '') {
    const headerLike = [];
    for (let r = partIiRow + 1; r < scanTo; r += 1) {
      const filled = (rows[r] || []).map((c) => norm(c)).filter(Boolean);
      if (!filled.length) continue;
      if (isColumnIndexRow(rows[r])) continue;
      headerLike.push(filled);
    }
    const band = headerLike[0] || [];
    if (band.length) {
      const yearTok = band.find((t) => /^(19|20)\d{2}$/.test(t));
      if (yearTok) values.year = values.year || yearTok;
      const lastNum = [...band].reverse().find((t) => /^-?\d+(\.\d+)?$/.test(t) && !/^(19|20)\d{2}$/.test(t));
      if (lastNum != null && values.balance === '') values.balance = lastNum;
    }
  }
  return values;
};

const padRows = (rows, minLen) => {
  const out = (Array.isArray(rows) ? rows : []).map((row) => [...row]);
  while (out.length < minLen) out.push(Array(FORM_F_KA_PDF_PART_I_LEAFS.length).fill(''));
  return out;
};

export const normalizeFormFKarnatakaPdfMatrix = (
  rows = [],
  colCount = 11,
  tableStartRow = 0,
  incomingMetaLines = []
) => {
  void tableStartRow;
  const src = [];
  (Array.isArray(incomingMetaLines) ? incomingMetaLines : []).forEach((line) => {
    const t = norm(line);
    if (t) src.push([t]);
  });
  (Array.isArray(rows) ? rows : []).forEach((row) => {
    src.push(Array.isArray(row) ? row.map((c) => norm(c)) : [norm(row)]);
  });
  const cols = Math.max(
    11,
    Number(colCount) || 11,
    ...src.map((row) => (Array.isArray(row) ? row.length : 0))
  );

  const identity = extractIdentity(src, cols);
  const partIRows = padRows(extractPartIRows(src, cols), MIN_PART_I_BODY_ROWS);
  const partIi = extractPartIi(src, cols);

  const model = {
    titles: [FORM_F_KA_PDF_TITLE, FORM_F_KA_PDF_RULE, FORM_F_KA_PDF_SUBTITLE],
    identity,
    partITitle: FORM_F_KA_PDF_PART_I_TITLE,
    partIRows,
    partIiTitle: FORM_F_KA_PDF_PART_II_TITLE,
    partIi,
  };

  const identityRows = [
    [
      '1',
      FORM_F_KA_PDF_IDENTITY_FIELDS[0].label,
      identity.slNo,
      '2',
      FORM_F_KA_PDF_IDENTITY_FIELDS[1].label,
      identity.dateOfEntry,
    ],
    [
      '3',
      FORM_F_KA_PDF_IDENTITY_FIELDS[2].label,
      identity.personName,
      '4',
      FORM_F_KA_PDF_IDENTITY_FIELDS[3].label,
      identity.fatherName,
    ],
  ];
  const groupRow = [];
  const leafRow = [];
  const numberRow = [];
  FORM_F_KA_PDF_PART_I_GROUPS.forEach((group) => {
    group.leafs.forEach((leaf, i) => {
      groupRow.push(i === 0 ? group.label : '');
      leafRow.push(leaf.label);
      numberRow.push(leaf.number);
    });
  });
  const tableRows = [groupRow, leafRow, numberRow, ...partIRows];
  const partIiHeader = ['Year', 'of credit', 'Availed', 'Balance'];
  const partIiNumbers = ['1', '2', '3', '4'];
  const partIiData = [partIi.year, partIi.credit, partIi.availed, partIi.balance];
  const partIiPad = Array.from({ length: MIN_PART_II_BODY_ROWS - 1 }, () => ['', '', '', '']);

  return {
    rows: [
      ...identityRows,
      [FORM_F_KA_PDF_PART_I_TITLE],
      ...tableRows,
      [FORM_F_KA_PDF_PART_II_TITLE],
      partIiHeader,
      partIiNumbers,
      partIiData,
      ...partIiPad,
    ],
    colCount: 11,
    tableStartRow: 3,
    metaLines: model.titles,
    formFKALayout: true,
    formFKAModel: model,
  };
};

export const applyFormFKarnatakaPdfNormalization = (target, normalized) => {
  if (!target || !normalized) return target;
  target.rows = normalized.rows;
  target.colCount = normalized.colCount;
  target.tableStartRow = normalized.tableStartRow;
  target.metaLines = Array.isArray(normalized.metaLines) ? normalized.metaLines : [];
  target.formFKALayout = true;
  target.formFKAModel = normalized.formFKAModel;
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
  const lineH = size + 1.5;
  const top = y + size;
  if (align === 'center') {
    doc.text(lines, x + w / 2, top, { align: 'center' });
  } else {
    doc.text(lines, x + 2, top);
  }
  return lines.length * lineH;
}

function colWidths(total, weights) {
  const sum = weights.reduce((a, b) => a + b, 0) || 1;
  return weights.map((w) => (total * w) / sum);
}

function ensurePage(doc, y, need, marginBottom = 28) {
  const pageHeight = doc.internal.pageSize.getHeight();
  if (y + need <= pageHeight - marginBottom) return y;
  doc.addPage();
  return 24;
}

function drawCenteredBand(doc, x, y, w, h, text, { bold = true, size = 11 } = {}) {
  strokeRect(doc, x, y, w, h);
  doc.setFont('helvetica', bold ? 'bold' : 'normal');
  doc.setFontSize(size);
  const lines = doc.splitTextToSize(String(text || ''), Math.max(12, w - 8));
  const blockH = lines.length * (size + 2);
  doc.text(lines, x + w / 2, y + (h - blockH) / 2 + size, { align: 'center' });
}

function drawIdentityPair(doc, x, y, w, h, left, right) {
  const numW = 22;
  const labelW = Math.min(168, Math.max(120, w * 0.42));
  const valueW = w - numW - labelW;
  const paint = (ox, field) => {
    strokeRect(doc, ox, y, numW, h);
    drawWrapped(doc, field.number, ox, y + (h - 11) / 2 - 2, numW, 10, 'center', true);
    strokeRect(doc, ox + numW, y, labelW, h);
    drawWrapped(doc, field.label, ox + numW, y + 3, labelW, 8, 'left', true);
    strokeRect(doc, ox + numW + labelW, y, valueW, h);
    drawWrapped(doc, field.value, ox + numW + labelW, y + 4, valueW, 9, 'left', false);
  };
  paint(x, left);
  paint(x + w, right);
}

function drawGroupedTable(doc, x, y, width, groups, dataRows, { groupH, leafH, numH, dataH }) {
  const leafs = [];
  groups.forEach((g) => g.leafs.forEach((leaf) => leafs.push({ ...leaf, parent: g.label })));
  const widths = colWidths(
    width,
    leafs.map((leaf) => leaf.weight || 8)
  );
  let cursor = x;
  groups.forEach((group) => {
    const spanW = group.leafs.reduce((sum, leaf) => {
      const idx = leafs.findIndex((l) => l.key === leaf.key);
      return sum + (widths[idx] || 0);
    }, 0);
    strokeRect(doc, cursor, y, spanW, groupH);
    if (group.label) drawWrapped(doc, group.label, cursor, y + 2, spanW, 7.5, 'center', true);
    cursor += spanW;
  });
  cursor = x;
  leafs.forEach((leaf, i) => {
    strokeRect(doc, cursor, y + groupH, widths[i], leafH);
    if (leaf.label) drawWrapped(doc, leaf.label, cursor, y + groupH + 2, widths[i], 7, 'center', true);
    cursor += widths[i];
  });
  cursor = x;
  leafs.forEach((leaf, i) => {
    strokeRect(doc, cursor, y + groupH + leafH, widths[i], numH);
    drawWrapped(doc, leaf.number, cursor, y + groupH + leafH + 1, widths[i], 8, 'center', true);
    cursor += widths[i];
  });
  const bodyY = y + groupH + leafH + numH;
  (dataRows || []).forEach((row, ri) => {
    cursor = x;
    leafs.forEach((leaf, i) => {
      strokeRect(doc, cursor, bodyY + ri * dataH, widths[i], dataH);
      const val = Array.isArray(row) ? row[i] : row?.[leaf.key] || '';
      drawWrapped(doc, val, cursor, bodyY + ri * dataH + 3, widths[i], 7.5, 'center', false);
      cursor += widths[i];
    });
  });
  return bodyY + (dataRows || []).length * dataH;
}

/**
 * Draw Form F as the Excel register: title band, 2×2 identity, PART I, PART II.
 */
export function drawFormFKarnatakaLeaveRegister(doc, matrix, startY = 22) {
  const pageWidth = doc.internal.pageSize.getWidth();
  const marginX = 22;
  const usableWidth = pageWidth - marginX * 2;
  const model =
    matrix?.formFKAModel ||
    normalizeFormFKarnatakaPdfMatrix(matrix?.rows || [], matrix?.colCount || 11, 0, matrix?.metaLines || [])
      .formFKAModel;
  let y = startY;
  const id = model.identity || {};

  drawCenteredBand(doc, marginX, y, usableWidth, 18, model.titles?.[0] || FORM_F_KA_PDF_TITLE, {
    bold: true,
    size: 12,
  });
  y += 18;
  drawCenteredBand(doc, marginX, y, usableWidth, 16, model.titles?.[1] || FORM_F_KA_PDF_RULE, {
    bold: false,
    size: 9,
  });
  y += 16;
  drawCenteredBand(doc, marginX, y, usableWidth, 18, model.titles?.[2] || FORM_F_KA_PDF_SUBTITLE, {
    bold: true,
    size: 11,
  });
  y += 18;

  const half = usableWidth / 2;
  const idH = 32;
  drawIdentityPair(
    doc,
    marginX,
    y,
    half,
    idH,
    { number: '1', label: FORM_F_KA_PDF_IDENTITY_FIELDS[0].label, value: id.slNo || '' },
    { number: '2', label: FORM_F_KA_PDF_IDENTITY_FIELDS[1].label, value: id.dateOfEntry || '' }
  );
  y += idH;
  drawIdentityPair(
    doc,
    marginX,
    y,
    half,
    idH,
    { number: '3', label: FORM_F_KA_PDF_IDENTITY_FIELDS[2].label, value: id.personName || '' },
    { number: '4', label: FORM_F_KA_PDF_IDENTITY_FIELDS[3].label, value: id.fatherName || '' }
  );
  y += idH;

  drawCenteredBand(doc, marginX, y, usableWidth, 16, model.partITitle || FORM_F_KA_PDF_PART_I_TITLE, {
    bold: true,
    size: 10,
  });
  y += 16;

  const partIRows = padRows(model.partIRows || [], MIN_PART_I_BODY_ROWS);
  y = drawGroupedTable(doc, marginX, y, usableWidth, FORM_F_KA_PDF_PART_I_GROUPS, partIRows, {
    groupH: 34,
    leafH: 18,
    numH: 14,
    dataH: 20,
  });

  y = ensurePage(doc, y + 10, 90);
  drawCenteredBand(doc, marginX, y, usableWidth, 16, model.partIiTitle || FORM_F_KA_PDF_PART_II_TITLE, {
    bold: true,
    size: 10,
  });
  y += 16;

  const partIiRows = padRows(
    [[model.partIi?.year || '', model.partIi?.credit || '', model.partIi?.availed || '', model.partIi?.balance || '']],
    MIN_PART_II_BODY_ROWS
  );
  const partIiWidth = usableWidth * 0.62;
  y = drawGroupedTable(doc, marginX, y, partIiWidth, FORM_F_KA_PDF_PART_II_GROUPS, partIiRows, {
    groupH: 22,
    leafH: 16,
    numH: 14,
    dataH: 16,
  });
  return y + 8;
}
