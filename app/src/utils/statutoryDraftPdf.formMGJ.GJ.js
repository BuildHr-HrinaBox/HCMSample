/**
 * Gujarat Form M — Identity Card (See Rule 16) PDF layout.
 * Excel is a sparse identity card (titles + photo box + label/value + signatures).
 * Generic grid conversion stacked every cell as a full-width table row; this
 * rebuilds the official card alignment instead.
 */

export const FORM_MGJ_GJ_PDF_TITLE = 'FORM - M';
export const FORM_MGJ_GJ_PDF_RULE = '(See Rule 16)';
export const FORM_MGJ_GJ_PDF_SUBTITLE = 'IDENTITY CARD';
export const FORM_MGJ_GJ_PDF_PHOTO_LABEL = 'Photograph';

export const FORM_MGJ_GJ_PDF_FIELDS = [
  {
    key: 'establishment',
    ordinal: 'a',
    label: '(a) Name and address of the establishment:',
    match: /name\s+and\s+address\s+of\s+the\s+establishment/i,
  },
  {
    key: 'workerNameAddress',
    ordinal: 'b',
    label: '(b) The full name and address of the worker:',
    match: /full\s+name\s+and\s+address\s+of\s+the\s+worker|name\s+and\s+address\s+of\s+the\s+worker/i,
  },
  {
    key: 'dateOfBirth',
    ordinal: 'c',
    label: '(c) Date of birth of the worker;',
    match: /date\s+of\s+birth\s+of\s+the\s+worker/i,
  },
  {
    key: 'dateOfJoining',
    ordinal: 'd',
    label: '(d) Date of joining the service in the establishment:',
    match: /date\s+of\s+joining\s+the\s+service/i,
  },
  {
    key: 'photograph',
    ordinal: 'e',
    label: '(e) Recent passport size photograph of the worker.',
    match: /passport\s+size\s+photograph|photograph\s+of\s+the\s+worker/i,
  },
  {
    key: 'contactNo',
    ordinal: 'f',
    label: '(f) Contact No.',
    match: /contact\s+no\.?/i,
  },
];

export const FORM_MGJ_GJ_PDF_SIGNATURES = [
  'Signature or left thumb impression of the worker.',
  'Signature of Manager or Authorized Agent.',
  'Date of Issue.',
];

const normalizeFormMGJText = (text) =>
  String(text || '')
    .replace(/\r\n/g, '\n')
    .replace(/\u00a0/g, ' ')
    .replace(/[ \t]+/g, ' ')
    .replace(/\n /g, '\n')
    .trim();

const lowerFormMGJText = (text) => normalizeFormMGJText(text).toLowerCase();

const cellAt = (row, c) => normalizeFormMGJText(row?.[c]);

const compactRowCells = (row) =>
  (Array.isArray(row) ? row : []).map((cell) => normalizeFormMGJText(cell)).filter(Boolean);

const looksLikeFormXIVEmploymentCardBlob = (blob) =>
  (/form\s*xiv\b/.test(blob) && /employment\s+card/.test(blob)) ||
  (/employment\s+card/.test(blob) && /see\s+rule\s*76/.test(blob));

export const looksLikeFormMGJGujaratPdfContext = (
  metaLines,
  rows,
  sheetName = '',
  fileName = ''
) => {
  const blob = [...(metaLines || []), ...(rows || []).slice(0, 36).flat(), sheetName || '', fileName || '']
    .join(' ')
    .toLowerCase();
  if (!blob || looksLikeFormXIVEmploymentCardBlob(blob)) return false;

  const hasFormM =
    /\bform[\s._-]*m\b/.test(blob) ||
    /form[\s._-]*m[\s._-]*gj/.test(blob) ||
    /\bform_m_gj\b/.test(blob);
  const hasIdentity = /identity\s+card/.test(blob);
  const hasPhotoWorker = /passport\s+size\s+photograph|photograph\s+of\s+the\s+worker/.test(blob);
  const hasEstablishment = /name\s+and\s+address\s+of\s+the\s+establishment/.test(blob);
  const hasWorker = /(?:full\s+)?name\s+and\s+address\s+of\s+the\s+worker/.test(blob);
  const hasGujarat = /gujarat|\b_gj\b|form_m_gj/.test(blob);
  const hasRule = /see\s+rule\s*(16|18)/.test(blob);

  if (hasFormM && hasIdentity) return true;
  if (hasFormM && hasGujarat && (hasPhotoWorker || hasWorker)) return true;
  if (hasIdentity && hasEstablishment && hasWorker && hasPhotoWorker) return true;
  if (hasFormM && hasRule && (hasPhotoWorker || hasWorker || hasEstablishment)) return true;
  return false;
};

const isFormMGJTitleText = (text) => {
  const t = normalizeFormMGJText(text);
  if (!t) return false;
  return (
    /^form\s*[-–]?\s*m\b/i.test(t) ||
    /^identity\s+card$/i.test(t) ||
    /^\(?\s*see\s+rule\s*(16|18)\s*\)?$/i.test(t)
  );
};

const isFormMGJPhotoBoxText = (text) => /^photograph$/i.test(normalizeFormMGJText(text));

const isFormMGJSignatureText = (text) => {
  const t = lowerFormMGJText(text);
  if (!t) return false;
  return (
    /signature\s+or\s+left\s+thumb/.test(t) ||
    /signature\s+of\s+manager/.test(t) ||
    /authorized\s+agent/.test(t) ||
    /^date\s+of\s+issue\.?$/.test(t)
  );
};

const isFormMGJIgnorableValue = (text) => {
  const t = normalizeFormMGJText(text);
  if (!t) return true;
  if (isFormMGJTitleText(t) || isFormMGJPhotoBoxText(t) || isFormMGJSignatureText(t)) return true;
  if (/^photograph$/i.test(t)) return true;
  if (/^enter\b/i.test(t)) return true;
  if (/^[_\s.\-:…]+$/u.test(t)) return true;
  return false;
};

const matchFormMGJFieldSpec = (text) => {
  const raw = normalizeFormMGJText(text);
  if (!raw) return null;
  for (let i = 0; i < FORM_MGJ_GJ_PDF_FIELDS.length; i += 1) {
    const spec = FORM_MGJ_GJ_PDF_FIELDS[i];
    if (spec.match.test(raw)) return spec;
    const ordinalRe = new RegExp(`^\\(?${spec.ordinal}\\)?(?:[\\.)]|\\s)`, 'i');
    if (ordinalRe.test(raw) && spec.match.test(raw.replace(/^\(?[a-f]\)?[\.\)]?\s*/i, ''))) {
      return spec;
    }
  }
  return null;
};

const peelValueFromLabelText = (text, spec) => {
  const raw = String(text || '').replace(/\s+/g, ' ').trim();
  if (!raw || !spec) return '';
  const match = raw.match(spec.match);
  if (!match || match.index == null) {
    const ordinalRe = new RegExp(`^\\(?${spec.ordinal}\\)?[\\.)]?\\s*`, 'i');
    if (!ordinalRe.test(raw)) return '';
  }
  const start = match && match.index != null ? match.index + match[0].length : raw.length;
  let rest = raw.slice(start).replace(/^[:.\-;,\s]+/, '').trim();
  rest = rest
    .replace(/^in\s+the\s+establ?ishment[:.\s]*/i, '')
    .replace(/^of\s+the\s+worker[:.\s]*/i, '')
    .replace(/^[-–]\s*/, '')
    .trim();
  if (!rest || isFormMGJIgnorableValue(rest) || matchFormMGJFieldSpec(rest)) return '';
  return rest;
};

const looksLikeFormMGJDateValue = (text) => {
  const t = normalizeFormMGJText(text);
  if (!t || t.length > 40) return false;
  return (
    /^\d{1,2}\s+(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*\.?\s+\d{2,4}$/i.test(t) ||
    /^\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4}$/.test(t) ||
    /^\d{4}-\d{2}-\d{2}$/.test(t)
  );
};

const looksLikeFormMGJPhoneValue = (text) => {
  const digits = String(text || '').replace(/\D/g, '');
  return digits.length >= 6 && digits.length <= 15;
};

const looksLikeFormMGJSiteName = (text) => {
  const t = normalizeFormMGJText(text);
  if (!t || t.length > 80) return false;
  return /\bsite\b/i.test(t) || /\b(limited|pvt|private|llp)\b/i.test(t);
};

const preferRicherValue = (current, next) => {
  const a = normalizeFormMGJText(current);
  const b = normalizeFormMGJText(next);
  if (!b) return a;
  if (!a) return b;
  if (a === b) return a;
  if (b.includes(a) && b.length > a.length) return b;
  if (a.includes(b)) return a;
  if (b.length > a.length + 8) return b;
  return a.includes('\n') ? a : [a, b].filter(Boolean).join('\n');
};

const collectScanRows = (rows, metaLines) => {
  const out = [];
  (Array.isArray(metaLines) ? metaLines : []).forEach((line) => {
    const text = normalizeFormMGJText(line);
    if (text) out.push([text]);
  });
  (Array.isArray(rows) ? rows : []).forEach((row) => {
    if (Array.isArray(row) && row.some((c) => String(c || '').trim())) out.push(row);
  });
  return out;
};

const takeAheadValue = (scanRows, fromIndex, spec) => {
  for (let i = fromIndex + 1; i < Math.min(scanRows.length, fromIndex + 4); i += 1) {
    const filled = compactRowCells(scanRows[i]);
    if (!filled.length) continue;
    if (filled.some((t) => isFormMGJTitleText(t) || isFormMGJSignatureText(t) || isFormMGJPhotoBoxText(t))) {
      return '';
    }
    if (filled.some((t) => matchFormMGJFieldSpec(t))) return '';
    const joined = filled.filter((t) => !isFormMGJIgnorableValue(t)).join('\n').trim();
    if (!joined) continue;
    if (spec.key === 'photograph') return '';
    return joined;
  }
  return '';
};

export const normalizeFormMGJGujaratPdfMatrix = (
  rows,
  colCount,
  tableStartRow = 0,
  metaLines = []
) => {
  void colCount;
  void tableStartRow;
  const scanRows = collectScanRows(rows, metaLines);
  const titles = [];
  const values = {
    establishment: '',
    workerNameAddress: '',
    dateOfBirth: '',
    dateOfJoining: '',
    photograph: '',
    contactNo: '',
  };
  const orphans = [];
  let ruleLine = '';

  const recordTitle = (text) => {
    if (/see\s+rule/i.test(text)) {
      ruleLine = /^\(/.test(text) ? text : `(${text.replace(/^\(|\)$/g, '')})`;
      return;
    }
    if (!titles.includes(text)) titles.push(text);
  };

  scanRows.forEach((row, r) => {
    const cells = Array.isArray(row) ? row : [];
    const filled = compactRowCells(cells);
    if (!filled.length) return;

    filled.forEach((text) => {
      if (isFormMGJTitleText(text)) recordTitle(text);
    });

    let labelCol = -1;
    let spec = null;
    for (let c = 0; c < cells.length; c += 1) {
      const t = cellAt(cells, c);
      if (!t) continue;
      const hit = matchFormMGJFieldSpec(t);
      if (hit) {
        labelCol = c;
        spec = hit;
        break;
      }
    }

    if (spec) {
      const labelText = cellAt(cells, labelCol);
      let value = peelValueFromLabelText(labelText, spec);
      if (!value) {
        const sameRow = [];
        for (let c = labelCol + 1; c < cells.length; c += 1) {
          const t = cellAt(cells, c);
          if (!t || isFormMGJIgnorableValue(t) || matchFormMGJFieldSpec(t)) continue;
          if (isFormMGJPhotoBoxText(t) || isFormMGJTitleText(t) || isFormMGJSignatureText(t)) {
            continue;
          }
          sameRow.push(t);
        }
        value = sameRow.join('\n').trim();
      }
      if (!value) value = takeAheadValue(scanRows, r, spec);
      if (spec.key !== 'photograph') {
        values[spec.key] = preferRicherValue(values[spec.key], value);
      }
      return;
    }

    filled.forEach((text) => {
      if (
        isFormMGJTitleText(text) ||
        isFormMGJPhotoBoxText(text) ||
        isFormMGJSignatureText(text) ||
        isFormMGJIgnorableValue(text)
      ) {
        return;
      }
      orphans.push(text);
    });
  });

  const unusedOrphans = orphans.filter((text) => {
    const t = normalizeFormMGJText(text);
    return !Object.values(values).some((v) => v && (v === t || String(v).includes(t)));
  });

  unusedOrphans.forEach((text) => {
    if (looksLikeFormMGJDateValue(text)) {
      if (!values.dateOfBirth) {
        values.dateOfBirth = text;
        return;
      }
      if (!values.dateOfJoining) {
        values.dateOfJoining = text;
        return;
      }
    }
    if (looksLikeFormMGJPhoneValue(text) && !values.contactNo) {
      values.contactNo = text;
      return;
    }
    if (!values.establishment) {
      values.establishment = text;
      return;
    }
    if (
      looksLikeFormMGJSiteName(text) &&
      !values.establishment.toLowerCase().includes(text.toLowerCase())
    ) {
      values.establishment = `${text}\n${values.establishment}`;
      return;
    }
    if (!values.workerNameAddress) {
      values.workerNameAddress = text;
    }
  });

  const orderedTitles = [];
  const formTitle = titles.find((t) => /^form\s*[-–]?\s*m\b/i.test(t)) || FORM_MGJ_GJ_PDF_TITLE;
  const identityTitle = titles.find((t) => /identity\s+card/i.test(t)) || FORM_MGJ_GJ_PDF_SUBTITLE;
  orderedTitles.push(formTitle);
  orderedTitles.push(ruleLine || FORM_MGJ_GJ_PDF_RULE);
  orderedTitles.push(identityTitle);

  const fieldRows = FORM_MGJ_GJ_PDF_FIELDS.map((spec) => [spec.label, values[spec.key] || '']);

  return {
    rows: fieldRows,
    colCount: 2,
    tableStartRow: 0,
    metaLines: orderedTitles,
    formMGJLayout: true,
    formMGJTitles: orderedTitles,
    formMGJFields: FORM_MGJ_GJ_PDF_FIELDS.map((spec) => ({
      key: spec.key,
      label: spec.label,
      value: values[spec.key] || '',
    })),
    formMGJSignatures: [...FORM_MGJ_GJ_PDF_SIGNATURES],
  };
};

export const applyFormMGJGujaratPdfNormalization = (target, normalized) => {
  if (!target || !normalized) return target;
  target.rows = normalized.rows;
  target.colCount = normalized.colCount;
  target.tableStartRow = normalized.tableStartRow;
  target.metaLines = Array.isArray(normalized.metaLines) ? normalized.metaLines : [];
  target.formMGJLayout = true;
  target.formMGJTitles = normalized.formMGJTitles || [];
  target.formMGJFields = normalized.formMGJFields || [];
  target.formMGJSignatures = normalized.formMGJSignatures || [...FORM_MGJ_GJ_PDF_SIGNATURES];
  return target;
};

const strokeRect = (doc, x, y, w, h) => {
  doc.setDrawColor(0, 0, 0);
  doc.setLineWidth(0.8);
  doc.rect(x, y, w, h, 'S');
};

/**
 * Draw Form M GJ as the Excel identity card: centered titles, photograph box
 * top-right, label/value on the same row, signatures on the right.
 */
export const drawFormMGJGujaratIdentityCard = (doc, matrix, startY = 36) => {
  const pageWidth = doc.internal.pageSize.getWidth();
  const marginX = 40;
  const usableWidth = pageWidth - marginX * 2;
  let y = startY;

  const titles = Array.isArray(matrix?.formMGJTitles) && matrix.formMGJTitles.length
    ? matrix.formMGJTitles
    : [FORM_MGJ_GJ_PDF_TITLE, FORM_MGJ_GJ_PDF_RULE, FORM_MGJ_GJ_PDF_SUBTITLE];
  const fields = Array.isArray(matrix?.formMGJFields) && matrix.formMGJFields.length
    ? matrix.formMGJFields
    : FORM_MGJ_GJ_PDF_FIELDS.map((spec, i) => ({
        key: spec.key,
        label: spec.label,
        value: String(matrix?.rows?.[i]?.[1] || ''),
      }));
  const signatures =
    Array.isArray(matrix?.formMGJSignatures) && matrix.formMGJSignatures.length
      ? matrix.formMGJSignatures
      : [...FORM_MGJ_GJ_PDF_SIGNATURES];

  const photoW = 92;
  const photoH = 110;
  const photoX = marginX + usableWidth - photoW;
  const titleAreaW = usableWidth - photoW - 16;
  const titleTop = y;

  titles.forEach((title, idx) => {
    const isFormName = /^form\s*[-–]?\s*m\b/i.test(title);
    const isIdentity = /identity\s+card/i.test(title);
    const size = isFormName || isIdentity ? 12 : 10;
    const bold = isFormName || isIdentity;
    doc.setFont('helvetica', bold ? 'bold' : 'normal');
    doc.setFontSize(size);
    doc.setTextColor(0, 0, 0);
    const lines = doc.splitTextToSize(String(title || ''), titleAreaW);
    const blockH = lines.length * (size + 3);
    doc.text(lines, marginX + titleAreaW / 2, y + size, { align: 'center' });
    y += blockH + 4;
  });

  strokeRect(doc, photoX, titleTop, photoW, photoH);
  doc.setFont('helvetica', 'normal');
  doc.setFontSize(9);
  doc.text(FORM_MGJ_GJ_PDF_PHOTO_LABEL, photoX + photoW / 2, titleTop + photoH / 2 + 3, {
    align: 'center',
  });

  y = Math.max(y, titleTop + photoH) + 22;

  const valueX = marginX + 292;
  const valueW = Math.max(80, marginX + usableWidth - valueX);

  fields.forEach((field) => {
    const label = String(field?.label || '').trim();
    const value = normalizeFormMGJText(field?.value || '');
    doc.setFont('helvetica', 'normal');
    doc.setFontSize(10);
    const labelW = Math.max(120, valueX - marginX - 8);
    const labelLines = doc.splitTextToSize(label, labelW);
    const valueLines = value ? doc.splitTextToSize(value, valueW) : [];
    const lineH = 13;
    const h = Math.max(labelLines.length, valueLines.length || 1) * lineH;
    doc.text(labelLines, marginX, y);
    if (valueLines.length) {
      doc.text(valueLines, valueX, y);
    }
    y += h + 16;
  });

  y += 18;
  doc.setFont('helvetica', 'normal');
  doc.setFontSize(10);
  signatures.forEach((line) => {
    const text = String(line || '').trim();
    if (!text) return;
    const wrapped = doc.splitTextToSize(text, usableWidth * 0.46);
    doc.text(wrapped, pageWidth - marginX, y, { align: 'right' });
    y += wrapped.length * 13 + 16;
  });

  return y;
};
