/**
 * Karnataka Form Q — Appointment Order (Rule 24(9A)).
 * Excel is a 2-column label | value sheet. Generic PDF meta splitting
 * otherwise stacks every cell as a full-width band.
 */

export const FORM_Q_KA_PDF_TITLE = 'FORM Q';
export const FORM_Q_KA_PDF_RULE = '(See Rule 24(9A))';
export const FORM_Q_KA_PDF_SUBTITLE = 'APPOINTMENT ORDER';

export const FORM_Q_KA_PDF_FIELDS = [
  {
    key: 'establishment',
    label: '1. Name & Address of the Establishment',
    match: /name\s*(?:&|and)\s*address\s+of\s+the\s+establishment/i,
  },
  {
    key: 'employer',
    label: '2. Name & Address of the Employer',
    match: /name\s*(?:&|and)\s*address\s+of\s+the\s+employer/i,
  },
  {
    key: 'employee',
    label: '3. Name of the Employee',
    match: /name\s+of\s+the\s+employee/i,
  },
  {
    key: 'postal',
    label: '4. His/Her Postal Address',
    match: /postal\s+address/i,
  },
  {
    key: 'permanent',
    label: '5. His/Her Permanent Address',
    match: /permanent\s+address/i,
  },
  {
    key: 'father',
    label: '6. Father/Husband Name',
    match: /father\s*\/?\s*husband/i,
  },
  {
    key: 'dob',
    label: '7. Date of Birth',
    match: /date\s+of\s+birth/i,
  },
  {
    key: 'entry',
    label: '8. Date of his/her entry into employment',
    match: /(?:date\s+of\s+his\/?her\s+)?entry\s+into\s+employment/i,
  },
  {
    key: 'designation',
    label: '9. Designation',
    match: /^(?:\d+[.)]\s*)?designation\b/i,
  },
  {
    key: 'nature',
    label: '10. Nature of work entrusted to him/her',
    match: /nature\s+of\s+work(?:\s+entrusted(?:\s+to\s+him\/?her)?)?/i,
  },
  {
    key: 'serial',
    label: '11. His/Her serial number in the Register of employment',
    match: /serial\s+number\s+in\s+the\s+register(?:\s+of\s+employment)?/i,
  },
  {
    key: 'wages',
    label: '12. Rates of wages payable to him/her',
    match: /rates?\s+of\s+wages\s+payable(?:\s+to\s+him\/?her)?/i,
  },
];

export const FORM_Q_KA_PDF_ACK =
  'Acknowledgement by the employer with date & signature';
export const FORM_Q_KA_PDF_SEAL = 'Seal of the Establishment';
export const FORM_Q_KA_PDF_SIGNATURE = 'Signature of employer';

const cellText = (value) =>
  String(value ?? '')
    .replace(/_x000d_/gi, '\n')
    .replace(/\u000d/g, '\n')
    .replace(/\r\n/g, '\n')
    .replace(/\r/g, '\n')
    .replace(/\n{3,}/g, '\n\n')
    .trim();

const flatText = (value) => cellText(value).replace(/\s+/g, ' ').trim();

const blobFrom = (metaLines, rows, sheetName = '', fileName = '') =>
  [...(metaLines || []), ...(rows || []).slice(0, 24).flat(), sheetName || '', fileName || '']
    .map((t) => flatText(t))
    .join(' ')
    .toLowerCase();

const isFormQFilename = (sheetName = '', fileName = '') => {
  const name = `${sheetName || ''} ${fileName || ''}`.toLowerCase();
  return (
    /form[\s._-]*q[\s._-]*ka\b/.test(name) ||
    /form[\s._-]*q[\s._-]*karnataka/.test(name) ||
    /form_q_karnataka/.test(name)
  );
};

const isMaharashtraMusterBlob = (blob) =>
  /full\s+name\s+of\s+the\s+worker/.test(blob) &&
  /date\s+of\s+the\s+month|working\s+hours|interval\s+for\s+rest/.test(blob);

const isGujaratAnnualReturnBlob = (blob) =>
  /annual\s+return/.test(blob) &&
  /gujarat/.test(blob) &&
  !/appointment\s+order/.test(blob) &&
  !/karnataka/.test(blob);

export const looksLikeFormQKarnatakaPdfContext = (
  metaLines = [],
  rows = [],
  sheetName = '',
  fileName = ''
) => {
  if (isFormQFilename(sheetName, fileName)) return true;
  const blob = blobFrom(metaLines, rows, sheetName, fileName);
  if (!blob) return false;
  if (isMaharashtraMusterBlob(blob)) return false;
  if (isGujaratAnnualReturnBlob(blob)) return false;

  const isFormQ = /(?:^|[^a-z0-9])form[\s._-]*q(?:[^a-z0-9]|$)/.test(blob);
  const isAppointment = /appointment\s+order/.test(blob);
  const isRule249A = /rule\s+24\s*\(\s*9\s*a/.test(blob);
  const isKA = /karnataka/.test(blob);
  const hasEmployment =
    /name\s+of\s+the\s+employee/.test(blob) &&
    (/rates?\s+of\s+wages\s+payable/.test(blob) ||
      /name\s*(?:&|and)\s*address\s+of\s+the\s+employer/.test(blob));

  if (isAppointment && (isFormQ || hasEmployment || isKA || isRule249A)) return true;
  if (isFormQ && isRule249A) return true;
  if (isFormQ && isKA && hasEmployment) return true;
  if (hasEmployment && isRule249A) return true;
  return false;
};

const isTitleLine = (text) => {
  const t = flatText(text);
  if (!t) return false;
  if (/^form\s*q\b/i.test(t) && t.length < 24) return true;
  if (/see\s+rule\s*24\s*\(\s*9\s*a/i.test(t)) return true;
  if (/^appointment\s+order$/i.test(t)) return true;
  return false;
};

const isWageRateLine = (text) => {
  const t = flatText(text);
  if (!t) return false;
  if (/rates?\s+of\s+wages\s+payable/i.test(t)) return false;
  return /^\s*[1-4][.)]\s*(basic|vda|other\s+allowances?|total)\b/i.test(t);
};

const isPlaceLabel = (text) => {
  const t = flatText(text);
  if (!t) return false;
  return /^place\s*:?\s*$/i.test(t) || /^place\s*:/i.test(t);
};

const isDateFooterLabel = (text) => {
  const t = flatText(text);
  if (!t) return false;
  if (/date\s+of\s+birth|entry\s+into\s+employment/i.test(t)) return false;
  return /^date\s*:?\s*$/i.test(t) || /^date\s*:/i.test(t);
};

const isSignatureLabel = (text) => /signature\s+of\s+employer/i.test(flatText(text));
const isAckLabel = (text) => /acknowledgements?\s+by\s+the\s+employer/i.test(flatText(text));
const isSealLabel = (text) => /seal\s+of\s+the\s+establishment/i.test(flatText(text));
const isSystemGeneratedDocumentNote = (text) =>
  /system\s+generated\s+document|computer\s+generated\s+document/i.test(flatText(text));

const matchFieldIndex = (text) => {
  const t = flatText(text);
  if (!t || isWageRateLine(t)) return -1;
  for (let i = 0; i < FORM_Q_KA_PDF_FIELDS.length; i += 1) {
    if (FORM_Q_KA_PDF_FIELDS[i].match.test(t)) return i;
  }
  return -1;
};

const isFieldLabel = (text) => matchFieldIndex(text) >= 0;

const isFooterLabel = (text) =>
  isPlaceLabel(text) ||
  isDateFooterLabel(text) ||
  isSignatureLabel(text) ||
  isAckLabel(text) ||
  isSealLabel(text);

const stripLabelPrefix = (text, spec) => {
  const src = cellText(text);
  if (!src || !spec) return '';
  const flat = flatText(src);
  const m = flat.match(spec.match);
  if (!m) return '';
  const cut = flat.slice(m.index + m[0].length).replace(/^[.\-:\s]+/, '').trim();
  if (!cut || isFieldLabel(cut) || isFooterLabel(cut) || isTitleLine(cut) || isWageRateLine(cut)) {
    return '';
  }
  if (
    /^(to|of|for|in|the|his\/?her|him\/?her|entrusted)\b/i.test(cut) &&
    cut.length < 40 &&
    !/\d/.test(cut)
  ) {
    return '';
  }
  return cut;
};

const valueAfterLabel = (text, kind) => {
  const src = flatText(text);
  if (!src) return '';
  const patterns = {
    place: /^place\s*:?\s*/i,
    date: /^date\s*:?\s*/i,
  };
  const re = patterns[kind];
  if (!re) return '';
  const rest = src.replace(re, '').replace(/^[.\-:\s]+/, '').trim();
  if (!rest || isFooterLabel(rest) || isFieldLabel(rest)) return '';
  return rest;
};

const normalizeWageRatesValue = (raw) => {
  const src = cellText(raw);
  if (!src) return '';
  const lines = src
    .split(/\n+/)
    .map((s) => s.replace(/[ \t]+/g, ' ').trim())
    .filter(Boolean);
  if (lines.length >= 3 && lines.some((l) => /basic/i.test(l)) && lines.some((l) => /total/i.test(l))) {
    return lines.join('\n');
  }
  const smashed = src.replace(/\s+/g, ' ').trim();
  const parts = smashed
    .split(/(?=\b[1-4][.)]\s*(?:Basic|VDA|Other\s+allowances?|Total)\b)/i)
    .map((s) => s.replace(/\s+/g, ' ').trim())
    .filter(Boolean);
  if (parts.length >= 3) return parts.join('\n');
  return src;
};

const appendValue = (current, next, { wages = false } = {}) => {
  const add = wages ? normalizeWageRatesValue(next) : cellText(next);
  if (!add) return current || '';
  if (!current) return add;
  if (wages) {
    const joined = `${current}\n${add}`;
    return normalizeWageRatesValue(joined);
  }
  if (flatText(current).toLowerCase() === flatText(add).toLowerCase()) return current;
  return current;
};

const nonEmptyCells = (row) => {
  const cells = Array.isArray(row) ? row : [row];
  const out = [];
  for (let c = 0; c < cells.length; c += 1) {
    const t = cellText(cells[c]);
    if (t) out.push({ c, t });
  }
  return out;
};

const canonicalTitle = (text) => {
  const t = flatText(text);
  if (/^form\s*q\b/i.test(t)) return FORM_Q_KA_PDF_TITLE;
  if (/see\s+rule\s*24\s*\(\s*9\s*a/i.test(t)) {
    return /\)$/.test(t) ? t : FORM_Q_KA_PDF_RULE;
  }
  if (/^appointment\s+order$/i.test(t)) return FORM_Q_KA_PDF_SUBTITLE;
  return t;
};

export const applyFormQKarnatakaPdfNormalization = (target, normalized) => {
  if (!target || !normalized) return target;
  target.rows = normalized.rows;
  target.colCount = normalized.colCount;
  target.tableStartRow = normalized.tableStartRow;
  if (Array.isArray(normalized.metaLines)) target.metaLines = normalized.metaLines;
  target.formQKALayout = true;
  return target;
};

export const normalizeFormQKarnatakaPdfMatrix = (
  rows = [],
  colCount = 2,
  tableStartRow = 0,
  incomingMetaLines = []
) => {
  void colCount;
  void tableStartRow;
  const srcRows = [];
  (Array.isArray(incomingMetaLines) ? incomingMetaLines : []).forEach((line) => {
    const t = cellText(line);
    if (t) srcRows.push([t]);
  });
  (Array.isArray(rows) ? rows : []).forEach((row) => {
    srcRows.push(Array.isArray(row) ? row : [row]);
  });

  const titles = [];
  const seenTitles = new Set();
  const pushTitle = (text) => {
    const line = canonicalTitle(text);
    if (!line) return;
    const key = line.toLowerCase();
    if (seenTitles.has(key)) return;
    seenTitles.add(key);
    titles.push(line);
  };

  const values = FORM_Q_KA_PDF_FIELDS.map(() => '');
  let place = '';
  let date = '';
  let signature = '';
  let acknowledgement = '';
  let seal = '';
  let pendingFieldIdx = -1;
  const preFieldOrphans = [];

  const assignField = (idx, val) => {
    if (idx < 0 || idx >= values.length) return;
    values[idx] = appendValue(values[idx], val, { wages: idx === 11 });
  };

  const takeSameRowValue = (cells, startIdx, { wages = false } = {}) => {
    const parts = [];
    for (let j = startIdx; j < cells.length; j += 1) {
      const t = cells[j].t;
      if (isTitleLine(t) || isFieldLabel(t) || isFooterLabel(t)) break;
      if (!wages && isWageRateLine(t) && startIdx !== j) break;
      parts.push(t);
    }
    if (!parts.length) return '';
    return wages ? normalizeWageRatesValue(parts.join('\n')) : parts.map((p) => cellText(p)).join(' ');
  };

  const splitPlaceDateCell = (text) => {
    const src = cellText(text);
    if (!src) return null;
    const hasPlace = /^place\s*:/i.test(flatText(src)) || /^place\s*:/im.test(src);
    const hasDate = /(^|\n)\s*date\s*:/i.test(src) || /\sdate\s*:/i.test(` ${flatText(src)}`);
    if (!hasPlace || !hasDate) return null;
    let p = '';
    let d = '';
    src.split(/\n+/).forEach((line) => {
      const t = line.trim();
      if (/^place\s*:/i.test(t)) p = valueAfterLabel(t, 'place') || p;
      else if (/^date\s*:/i.test(t) && !/birth|entry/i.test(t)) d = valueAfterLabel(t, 'date') || d;
    });
    return { place: p, date: d };
  };

  srcRows.forEach((row) => {
    const cells = nonEmptyCells(row);
    if (!cells.length) return;
    const joined = cells.map((x) => x.t).join(' ');
    if (isSystemGeneratedDocumentNote(joined)) return;

    let used = false;
    for (let i = 0; i < cells.length; i += 1) {
      const t = cells[i].t;
      const placeDate = splitPlaceDateCell(t);
      if (placeDate) {
        pendingFieldIdx = -1;
        if (placeDate.place) place = place || placeDate.place;
        if (placeDate.date) date = date || placeDate.date;
        const beside = takeSameRowValue(cells, i + 1);
        if (beside && isSignatureLabel(beside)) signature = signature || FORM_Q_KA_PDF_SIGNATURE;
        used = true;
        continue;
      }
      if (isTitleLine(t)) {
        pushTitle(t);
        used = true;
        continue;
      }
      const fieldIdx = matchFieldIndex(t);
      if (fieldIdx >= 0) {
        const inline = stripLabelPrefix(t, FORM_Q_KA_PDF_FIELDS[fieldIdx]);
        const beside = takeSameRowValue(cells, i + 1, { wages: fieldIdx === 11 });
        assignField(fieldIdx, beside || inline);
        pendingFieldIdx = beside ? -1 : fieldIdx;
        used = true;
        if (beside) break;
        continue;
      }
      if (isWageRateLine(t)) {
        assignField(11, takeSameRowValue(cells, i, { wages: true }) || t);
        pendingFieldIdx = 11;
        used = true;
        break;
      }
      if (isPlaceLabel(t)) {
        pendingFieldIdx = -1;
        const inline = valueAfterLabel(t, 'place');
        const beside = takeSameRowValue(cells, i + 1);
        if (inline) place = place || inline;
        else if (beside && !isSignatureLabel(beside)) place = place || beside;
        if (beside && isSignatureLabel(beside)) signature = signature || FORM_Q_KA_PDF_SIGNATURE;
        used = true;
        continue;
      }
      if (isDateFooterLabel(t)) {
        const inline = valueAfterLabel(t, 'date');
        const beside = takeSameRowValue(cells, i + 1);
        if (inline) date = date || inline;
        else if (beside && !isSignatureLabel(beside)) date = date || beside;
        if (beside && isSignatureLabel(beside)) signature = signature || FORM_Q_KA_PDF_SIGNATURE;
        used = true;
        continue;
      }
      if (isSignatureLabel(t)) {
        signature = FORM_Q_KA_PDF_SIGNATURE;
        used = true;
        continue;
      }
      if (isAckLabel(t)) {
        acknowledgement = FORM_Q_KA_PDF_ACK;
        const beside = takeSameRowValue(cells, i + 1);
        if (beside && isSealLabel(beside)) seal = FORM_Q_KA_PDF_SEAL;
        used = true;
        continue;
      }
      if (isSealLabel(t)) {
        seal = FORM_Q_KA_PDF_SEAL;
        used = true;
        continue;
      }
    }

    if (!used) {
      const text = cells.map((cell) => cell.t).join(' ');
      if (pendingFieldIdx >= 0) {
        assignField(pendingFieldIdx, pendingFieldIdx === 11 ? normalizeWageRatesValue(text) : text);
        if (pendingFieldIdx !== 11) pendingFieldIdx = -1;
      } else {
        preFieldOrphans.push(text);
      }
    }
  });

  preFieldOrphans.forEach((orphan) => {
    if (!orphan) return;
    if (!values[1]) assignField(1, orphan);
    else if (!values[0]) assignField(0, orphan);
  });

  if (!titles.some((t) => /^form\s*q\b/i.test(t))) titles.unshift(FORM_Q_KA_PDF_TITLE);
  if (!titles.some((t) => /see\s+rule\s*24/i.test(t))) {
    const formIdx = titles.findIndex((t) => /^form\s*q\b/i.test(t));
    titles.splice(formIdx + 1, 0, FORM_Q_KA_PDF_RULE);
  }
  if (!titles.some((t) => /appointment\s+order/i.test(t))) titles.push(FORM_Q_KA_PDF_SUBTITLE);

  const bodyRows = FORM_Q_KA_PDF_FIELDS.map((spec, idx) => [
    spec.label,
    idx === 11 ? normalizeWageRatesValue(values[idx]) : cellText(values[idx]),
  ]);

  const placeLine = place ? `Place: ${place}` : 'Place:';
  const dateLine = date ? `Date: ${date}` : 'Date:';
  const footerRows = [
    [`${placeLine}\n${dateLine}`, signature || FORM_Q_KA_PDF_SIGNATURE],
    [acknowledgement || FORM_Q_KA_PDF_ACK, seal || FORM_Q_KA_PDF_SEAL],
  ];

  return {
    rows: [...bodyRows, ...footerRows],
    colCount: 2,
    tableStartRow: 0,
    metaLines: titles,
    formQKALayout: true,
  };
};
