/** Form XXIII Gujarat — Register of Overtime PDF helpers (12-col OT band). */

export const FORM_XXIII_GJ_OT_PDF_HEADERS = [
  'Serial No.',
  'Name and surname of workmen',
  "Father's/Husband's name",
  'Sex',
  'Designation/ Nature of employment',
  'Dates on which overtime worked',
  'Total overtime worked or production in case of piece rates',
  'Normal rate of wages',
  'Overtime rate of wages',
  'Overtime earnings',
  'Date on which overtime wages paid',
  'Remarks',
];

const normalizeGjXxiiiText = (text) =>
  String(text || '')
    .replace(/\s+/g, ' ')
    .trim();

const lowerGjXxiiiText = (text) => normalizeGjXxiiiText(text).toLowerCase();

const compactGjCompare = (text) =>
  lowerGjXxiiiText(text)
    .replace(/[\u2018\u2019\u2032\u0060']/g, '')
    .replace(/[^a-z0-9 ]/g, '')
    .trim();

const isGjOtSerialHeader = (header) =>
  /(?:^|\s)(?:sr|s|serial)\.?\s*no\b/.test(lowerGjXxiiiText(header));

/** Official 1–11 index under OT headers; Remarks column stays blank (no "12"). */
export const buildFormXXIIIGJOtPdfColumnIndexRow = () =>
  FORM_XXIII_GJ_OT_PDF_HEADERS.map((header, i) => {
    if (/remarks?/.test(lowerGjXxiiiText(header))) return '';
    return String(i + 1);
  });

/** Strip thead / index-row text accidentally stored on employee rows. */
export const sanitizeFormXXIIIGJOtMatrixCell = (value, header, columnIndex = -1) => {
  const s = String(value ?? '').trim();
  if (!s) return '';
  const cellNorm = compactGjCompare(s);
  const hdrNorm = compactGjCompare(header);
  if (cellNorm && hdrNorm && cellNorm === hdrNorm) return '';
  if (/^remarks?$/.test(cellNorm) && /remarks?/.test(hdrNorm)) return '';
  const colIdx = Number(columnIndex);
  if (colIdx >= 0 && !isGjOtSerialHeader(header) && /^\d{1,2}$/.test(s) && Number(s) === colIdx + 1) {
    return '';
  }
  return s;
};

export const scrubFormXXIIIGJOtPdfDataLine = (line) =>
  FORM_XXIII_GJ_OT_PDF_HEADERS.map((header, j) =>
    sanitizeFormXXIIIGJOtMatrixCell(Array.isArray(line) ? line[j] : '', header, j)
  );

/** Rows that are only template header/index bleed (often in Remarks / col 12). */
export const isFormXXIIIGJOtSpuriousPdfDataRow = (line) => {
  const scrubbed = scrubFormXXIIIGJOtPdfDataLine(line);
  const name = String(scrubbed[1] || '').trim();
  const nonEmpty = scrubbed
    .map((cell, idx) => ({ cell: String(cell || '').trim(), idx }))
    .filter((x) => x.cell);
  if (nonEmpty.length === 0) return true;
  if (nonEmpty.length === 1 && nonEmpty[0].idx === 11) return true;
  if (!name && /^remarks?$/i.test(String(scrubbed[11] || ''))) return true;
  if (
    !name &&
    nonEmpty.length <= 3 &&
    nonEmpty.every(
      (x) =>
        x.idx === 11 ||
        x.idx === 0 ||
        /^remarks?$/i.test(x.cell) ||
        /^\d{1,2}$/.test(x.cell)
    )
  ) {
    return true;
  }
  return false;
};

const joinUniqueRowCells = (row = []) => {
  const unique = [];
  (Array.isArray(row) ? row : []).forEach((cell) => {
    const text = normalizeGjXxiiiText(cell);
    if (text && !unique.includes(text)) unique.push(text);
  });
  return unique.join(' ').trim();
};

export const looksLikeFormXXIIIGJOtPdfContext = (metaLines = [], rows = [], sheetName = '', fileName = '') => {
  const blob = [
    ...(metaLines || []),
    ...(rows || []).slice(0, 28).flat(),
    sheetName || '',
    fileName || '',
  ]
    .map(lowerGjXxiiiText)
    .join(' ');
  if (!blob) return false;
  const hasForm = /form[\s._-]*xxiii(?![a-z])/.test(blob);
  const hasOvertime = /register[\s._-]*of[\s._-]*overtime/.test(blob);
  const hasGujarat =
    /gujarat/.test(blob) ||
    /form[\s._-]*xxiii[\s._-]*gj/.test(blob) ||
    /\bxxiii[_\s.-]*gj\b/.test(blob) ||
    /form_xxiii_gj/.test(blob) ||
    (/\bgj\b/.test(blob) && hasForm);
  return hasForm && hasOvertime && hasGujarat;
};

/** Keep exactly the August 12 OT columns (drop leftover wages M–R empties). */
export const trimFormXXIIIGJOtPdfToTwelveColumns = (rows, colCount) => {
  const list = Array.isArray(rows) ? rows : [];
  const target = 12;
  const nextRows = list.map((row) => {
    const line = Array.isArray(row) ? row.slice(0, target) : [];
    while (line.length < target) line.push('');
    return line;
  });
  return {
    rows: nextRows,
    colCount: Math.min(Math.max(1, Number(colCount) || target), target),
  };
};

export const formXXIIIGJOtPdfColumnWeight = (headerText = '', maxDataLen = 0) => {
  const lower = lowerGjXxiiiText(headerText);
  const dataLen = Math.max(0, Number(maxDataLen) || 0);
  if (/(?:^|\s)(?:sr|s|serial)\.?\s*no\b/.test(lower)) return 5.5;
  if (/^(sex|gender)$/.test(lower)) return 5;
  if (/name and surname|name of workman|father|husband/.test(lower)) {
    return Math.max(11, Math.min(16, 10 + Math.min(dataLen, 18) * 0.25));
  }
  if (/designation|nature of employment/.test(lower)) {
    return Math.max(9, Math.min(13, 8 + Math.min(dataLen, 14) * 0.2));
  }
  if (/dates?\s+on\s+which\s+overtime|total\s+overtime/.test(lower)) return 8.5;
  if (/normal\s+rate|overtime\s+rate|overtime\s+earnings/.test(lower)) return 8;
  if (/date\s+on\s+which\s+overtime|wages\s+paid/.test(lower)) return 8;
  if (/remarks?/.test(lower)) return 7;
  return 0;
};

/** FORM XXIII / See rule / Register of Overtime title band. */
export const isFormXXIIIGJOtTitleRow = (row = []) => {
  const blob = lowerGjXxiiiText(joinUniqueRowCells(row));
  if (!blob) return false;
  if (/name and address|nature and location|principal employer/.test(blob)) return false;
  return (
    /form[\s._-]*xxiii(?![a-z])/.test(blob) ||
    /register[\s._-]*of[\s._-]*overtime/.test(blob) ||
    /see\s+rule\s*78/.test(blob)
  );
};

/** Contractor / nature / establishment / principal employer rows. */
export const isFormXXIIIGJOtAdminRow = (row = []) => {
  const blob = lowerGjXxiiiText(joinUniqueRowCells(row));
  if (!blob) return false;
  return (
    /name and address of contractor/.test(blob) ||
    /nature and location of work/.test(blob) ||
    /name and address of (?:the )?establishment/.test(blob) ||
    /name and address of principal employer/.test(blob) ||
    (/principal\s+employer/.test(blob) && /name and address/.test(blob))
  );
};

/** Form XIV-style ordinal preamble — never the OT table. */
export const isFormXXIIIGJOtPreambleRow = (row = []) => {
  const raw = joinUniqueRowCells(row);
  const blob = lowerGjXxiiiText(raw);
  if (!blob) return false;
  if (/^\d+[\.\)]\s+/.test(blob) && /name of the workman|wage period|tenure of employment|nature of employment|wage rate/.test(blob)) {
    return true;
  }
  if (blob.includes('serial') && blob.includes('register') && (blob.includes('workmen') || blob.includes('workman'))) {
    return true;
  }
  if (/\.{5,}|_{5,}|…/.test(raw)) {
    return (
      /name of the workman/.test(blob) ||
      /nature of employment/.test(blob) ||
      (/wage rate/.test(blob) && /piece/.test(blob)) ||
      /wage period/.test(blob) ||
      /tenure of employment/.test(blob)
    );
  }
  return false;
};

/** August 12-col OT thead (Serial No. + surname/workman + overtime). */
export const isFormXXIIIGJOtTableHeaderRow = (row = []) => {
  const blob = lowerGjXxiiiText(joinUniqueRowCells(row));
  if (!blob || isFormXXIIIGJOtPreambleRow(row)) return false;
  const hasSerial = /(?:^|\s)(?:sr|s|serial)\.?\s*no\b/.test(blob);
  const hasName =
    /name and surname of workmen/.test(blob) ||
    /name of workman/.test(blob) ||
    /surname of workmen/.test(blob);
  const hasOt =
    /\bovertime\b/.test(blob) ||
    /normal\s+rate\s+of\s+wages/.test(blob) ||
    /dates?\s+on\s+which\s+overtime/.test(blob);
  return hasSerial && (hasName || hasOt) && (hasOt || /father|husband|sex|designation/.test(blob));
};

export const isFormXXIIIGJOtNumberRow = (row = []) => {
  const filled = (Array.isArray(row) ? row : [])
    .map((cell) => String(cell || '').trim())
    .filter(Boolean);
  if (filled.length < 8) return false;
  const nums = filled.filter((cell) => /^\d{1,2}$/.test(cell)).map((cell) => Number(cell));
  if (nums.length < 8) return false;
  if (nums[0] !== 1) return false;
  for (let i = 1; i < nums.length; i += 1) {
    if (nums[i] !== nums[i - 1] + 1) return false;
  }
  return nums[nums.length - 1] >= 10;
};

export const isFormXXIIIGJOtSignatureRow = (row = []) =>
  /signature\s+of\s+contractor/i.test(joinUniqueRowCells(row));

export const isFormXXIIIGJOtCenterValueHeader = (headerText = '') => {
  const lower = lowerGjXxiiiText(headerText);
  if (!lower) return false;
  if (/(?:^|\s)(?:sr|s|serial)\.?\s*no\b/.test(lower)) return true;
  if (/^(sex|gender)$/.test(lower) || /\bsex\b/.test(lower)) return true;
  if (/dates?\s+on\s+which\s+overtime/.test(lower)) return true;
  if (/total\s+overtime/.test(lower)) return true;
  if (/normal\s+rate/.test(lower)) return true;
  if (/overtime\s+rate/.test(lower)) return true;
  if (/overtime\s+earnings/.test(lower)) return true;
  if (/date\s+on\s+which\s+overtime|wages\s+paid/.test(lower)) return true;
  return false;
};

const pushMetaUnique = (list, seen, raw) => {
  const text = normalizeGjXxiiiText(raw);
  if (!text) return;
  if (/^this\s+is\s+a\s+system\s+generated\s+document$/i.test(text)) return;
  const key = lowerGjXxiiiText(text);
  if (seen.has(key)) return;
  seen.add(key);
  list.push(text);
};

const expandMetaSegments = (text) => {
  const line = normalizeGjXxiiiText(text);
  if (!line) return [];
  // Keep label:value admin lines intact (do not split on colon).
  if (
    /name and address|nature and location|principal employer/i.test(line) &&
    /:/.test(line)
  ) {
    return [line];
  }
  return [line];
};

export const getFormXXIIIGJOtHeaderTitles = (metaLines = [], rows = [], tableStart = 0) => {
  const pool = [
    ...(metaLines || []),
    ...(rows || []).slice(0, Math.max(0, Number(tableStart) || 0)).flat(),
  ]
    .map(normalizeGjXxiiiText)
    .filter(Boolean);

  const form =
    pool.find((l) => /^form\s*xxiii\b/i.test(l) && l.length < 40) ||
    pool.find((l) => /form[\s._-]*xxiii/i.test(l) && l.length < 48) ||
    'FORM XXIII';
  const rule =
    pool.find((l) => /see\s+rule\s*78/i.test(l)) || '1[See rule 78 (1) (a) (iii)]';
  const register =
    pool.find((l) => /register\s+of\s+overtime/i.test(l)) || 'Register of Overtime';
  return [form, rule, register].map(normalizeGjXxiiiText).filter(Boolean);
};

/**
 * Rebuild matrix to match the August PDF model:
 * - titles + admin lines in meta (full-width bands)
 * - 12-col OT thead + 1..12 index + employee rows only
 * - drop Form XIV preamble / signature / wages leftovers
 */
export const normalizeFormXXIIIGJOtPdfMatrix = (
  rows,
  colCount,
  tableStartRow = 0,
  metaLines = []
) => {
  const src = Array.isArray(rows) ? rows : [];
  const target = 12;
  const metaOut = [];
  const seen = new Set();

  (metaLines || []).forEach((line) => {
    expandMetaSegments(line).forEach((seg) => pushMetaUnique(metaOut, seen, seg));
  });

  let headerIdx = -1;
  for (let r = 0; r < src.length; r += 1) {
    if (isFormXXIIIGJOtTableHeaderRow(src[r])) {
      headerIdx = r;
      break;
    }
  }
  if (headerIdx < 0) {
    // Fallback: first non-preamble wage-register-looking serial row after admin.
    for (let r = Math.max(0, Number(tableStartRow) || 0); r < src.length; r += 1) {
      const blob = lowerGjXxiiiText(joinUniqueRowCells(src[r]));
      if (!blob || isFormXXIIIGJOtPreambleRow(src[r])) continue;
      if (/(?:^|\s)(?:sr|s|serial)\.?\s*no\b/.test(blob) && /name|sex|overtime|designation/.test(blob)) {
        headerIdx = r;
        break;
      }
    }
  }

  const scanEnd = headerIdx >= 0 ? headerIdx : Math.min(src.length, Math.max(Number(tableStartRow) || 0, 16));
  for (let r = 0; r < scanEnd; r += 1) {
    const row = src[r] || [];
    if (isFormXXIIIGJOtPreambleRow(row)) continue;
    if (isFormXXIIIGJOtSignatureRow(row)) continue;
    const joined = joinUniqueRowCells(row);
    if (!joined) continue;
    if (isFormXXIIIGJOtTitleRow(row) || isFormXXIIIGJOtAdminRow(row) || joined.length > 8) {
      expandMetaSegments(joined).forEach((seg) => pushMetaUnique(metaOut, seen, seg));
    }
  }

  // Ensure canonical title order when titles were found (or inject defaults).
  const titles = getFormXXIIIGJOtHeaderTitles(metaOut, [], 0);
  const withoutTitles = metaOut.filter(
    (line) =>
      !/^form\s*xxiii\b/i.test(line) &&
      !/see\s+rule\s*78/i.test(line) &&
      !/register\s+of\s+overtime/i.test(line)
  );
  const adminOrder = [
    /name and address of contractor/i,
    /nature and location of work/i,
    /name and address of (?:the )?establishment/i,
    /principal employer/i,
  ];
  const orderedAdmin = [];
  adminOrder.forEach((re) => {
    withoutTitles.forEach((line) => {
      if (re.test(line) && !orderedAdmin.includes(line)) orderedAdmin.push(line);
    });
  });
  withoutTitles.forEach((line) => {
    if (!orderedAdmin.includes(line)) orderedAdmin.push(line);
  });
  const finalMeta = [...titles, ...orderedAdmin];

  const pad12 = (row) => {
    const line = (Array.isArray(row) ? row : []).slice(0, target).map((c) => normalizeGjXxiiiText(c));
    while (line.length < target) line.push('');
    return line;
  };

  const tableRows = [];
  tableRows.push(FORM_XXIII_GJ_OT_PDF_HEADERS.slice());
  tableRows.push(buildFormXXIIIGJOtPdfColumnIndexRow());

  const dataStart =
    headerIdx >= 0
      ? headerIdx + (isFormXXIIIGJOtNumberRow(src[headerIdx + 1] || []) ? 2 : 1)
      : Math.max(0, Number(tableStartRow) || 0);

  let employeeRowsAdded = 0;
  for (let r = dataStart; r < src.length; r += 1) {
    const row = src[r] || [];
    if (isFormXXIIIGJOtSignatureRow(row)) break;
    if (isFormXXIIIGJOtPreambleRow(row)) continue;
    if (isFormXXIIIGJOtTitleRow(row) || isFormXXIIIGJOtAdminRow(row)) continue;
    if (isFormXXIIIGJOtTableHeaderRow(row) || isFormXXIIIGJOtNumberRow(row)) continue;
    const line = scrubFormXXIIIGJOtPdfDataLine(pad12(row));
    if (isFormXXIIIGJOtSpuriousPdfDataRow(line)) {
      if (employeeRowsAdded > 0) break;
      continue;
    }
    const hasData = line.some((cell, idx) => idx > 0 && String(cell || '').trim());
    if (!hasData) continue;
    // Skip pure system-note rows.
    if (/^this\s+is\s+a\s+system\s+generated\s+document$/i.test(joinUniqueRowCells(line))) continue;
    tableRows.push(line);
    employeeRowsAdded += 1;
  }

  return {
    rows: tableRows,
    colCount: target,
    tableStartRow: 0,
    metaLines: finalMeta,
  };
};
