import { jsPDF } from 'jspdf';

const FORM_XV_AP_WORKMAN_LABELS = [
  'Name and address of the workman',
  'Age or date of Birth',
  'Identification Marks',
  "Father's / Husband's Name"
];

const FORM_XV_AP_TABLE_HEADERS = [
  'Sr.No.',
  'From',
  'To',
  'Nature of work done',
  'Rate of wage (with particulars of work)',
  'Remarks'
];

const isEmptyOrPlaceholder = (value) => {
  const s = String(value ?? '').trim();
  if (!s) return true;
  const lower = s.toLowerCase();
  return /^enter\b/.test(lower) || lower.includes('enter ') || lower.includes('select ');
};

const normalizeHeaderKey = (h) =>
  String(h || '')
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();

const pickRowValue = (row, predicate) => {
  if (!row || typeof row !== 'object') return '';
  const keys = Object.keys(row);
  for (let i = 0; i < keys.length; i += 1) {
    const key = keys[i];
    if (!predicate(key)) continue;
    const value = String(row[key] ?? '').trim();
    if (value && !isEmptyOrPlaceholder(value)) return value;
  }
  return '';
};

export const looksLikeFormXVAPPdfContext = (metaLines = [], rows = [], sheetName = '') => {
  const blob = [...(metaLines || []), ...(rows || []).slice(0, 40).flat(), sheetName || '']
    .join(' ')
    .toLowerCase();
  return (
    /form\s*xv\b|form[\s._-]*xv[\s._-]/.test(blob) &&
    /service\s+certificate/.test(blob) &&
    (/central\s*(?:&|and)\s*a\.?\s*p\.?|central\/\s*a\.?\s*p\.?|andhra\s+pradesh|\ba\.?p\.?\s+rules/.test(blob) ||
      /name\s+and\s+address\s+of\s+the\s+workman/.test(blob))
  );
};

export const formXVAPRowHasCorrespondingEmployee = (row) => {
  if (!row || typeof row !== 'object') return false;
  const workman = pickRowValue(
    row,
    (h) => {
      const n = normalizeHeaderKey(h);
      return (n.includes('workman') || n.includes('workmen')) && (n.includes('name') || n.includes('address'));
    }
  );
  return Boolean(workman);
};

export const filterFormXVAPCorrespondingEmployeeRows = (rows) =>
  (Array.isArray(rows) ? rows : []).filter((row) => formXVAPRowHasCorrespondingEmployee(row));

const pickHeaderValue = (headerFormData = {}, parsedFormHeader = {}, key, labelMatch) => {
  const direct = String(headerFormData?.[key] ?? '').trim();
  if (direct && !isEmptyOrPlaceholder(direct)) return direct;
  const fields = Array.isArray(parsedFormHeader?.fields) ? parsedFormHeader.fields : [];
  const field = fields.find((f) => {
    if (key && f?.key === key) return true;
    const lbl = String(f?.label || '');
    return labelMatch && labelMatch.test(lbl);
  });
  const fromField = String(field?.value ?? headerFormData?.[field?.key] ?? '').trim();
  return fromField && !isEmptyOrPlaceholder(fromField) ? fromField : '';
};

export const buildFormXVAPCertificatePageModel = ({
  headerFormData = {},
  parsedFormHeader = {},
  row = {},
  fileName = ''
} = {}) => {
  const titleFromHeader = String(parsedFormHeader?.title || '').trim();
  const subtitle = String(parsedFormHeader?.subtitle || '').trim();
  const reference = String(parsedFormHeader?.reference || '').trim();
  const titles = [];
  if (/form[\s._-]*xv/i.test(titleFromHeader) || /form[\s._-]*xv/i.test(fileName)) {
    titles.push('FORM XV');
  } else {
    titles.push(titleFromHeader || 'FORM XV');
  }
  titles.push(subtitle && /service\s+certificate/i.test(subtitle) ? subtitle : 'SERVICE CERTIFICATE');
  if (reference) {
    titles.push(reference);
  } else if (/andhra|a\.?\s*p\.?/i.test(`${titleFromHeader} ${fileName} ${reference}`)) {
    titles.push('[Vide rule 77 of Contract Labour (Regul. & Abolition) Central & A.P. Rules, 1971]');
  } else {
    titles.push('[Vide rule 77 of Contract Labour (Regulation & Abolition) Rules, 1971]');
  }

  const contractor = pickHeaderValue(
    headerFormData,
    parsedFormHeader,
    'form_xv_contractor',
    /name\s+and\s+address\s+of\s+(?:the\s+)?contractor/i
  );
  const establishment = pickHeaderValue(
    headerFormData,
    parsedFormHeader,
    'form_xv_establishment_contract_carried',
    /establishment/i
  );
  const nature = pickHeaderValue(
    headerFormData,
    parsedFormHeader,
    'form_xv_nature_location_work',
    /nature\s+and\s+location/i
  );
  const principal = pickHeaderValue(
    headerFormData,
    parsedFormHeader,
    'form_xv_principal_employer',
    /principal\s+employer/i
  );

  const workman = pickRowValue(row, (h) => {
    const n = normalizeHeaderKey(h);
    return (n.includes('workman') || n.includes('workmen')) && (n.includes('name') || n.includes('address'));
  });
  const age = pickRowValue(row, (h) => {
    const n = normalizeHeaderKey(h);
    return (n.includes('age') && n.includes('birth')) || n === 'age';
  });
  const marks = pickRowValue(row, (h) => {
    const n = normalizeHeaderKey(h);
    return n.includes('identification') && n.includes('mark');
  });
  const father = pickRowValue(row, (h) => {
    const n = normalizeHeaderKey(h);
    return (n.includes('father') || n.includes('husband')) && n.includes('name');
  });
  const from = pickRowValue(row, (h) => {
    const n = normalizeHeaderKey(h);
    return n.includes('from') && (n.includes('period') || n.includes('employed'));
  });
  const to = pickRowValue(row, (h) => {
    const n = normalizeHeaderKey(h);
    return n.includes('to') && !n.includes('from') && (n.includes('period') || n.includes('employed'));
  });
  const natureOfWork = pickRowValue(row, (h) => {
    const n = normalizeHeaderKey(h);
    return n.includes('nature') && n.includes('work') && !n.includes('location');
  });
  const rate = pickRowValue(row, (h) => {
    const n = normalizeHeaderKey(h);
    return n.includes('rate') && n.includes('wage');
  });
  const remarks = pickRowValue(row, (h) => normalizeHeaderKey(h).includes('remarks'));

  return {
    titles: titles.filter(Boolean),
    adminRows: [
      [contractor, establishment],
      [nature, principal]
    ],
    adminLabels: [
      ['Name and address of contractor', 'Name and address of establishment in/under which contract is carried on'],
      ['Nature and location of work', 'Name and address of principal employer']
    ],
    fieldRows: [
      [FORM_XV_AP_WORKMAN_LABELS[0], workman],
      [FORM_XV_AP_WORKMAN_LABELS[1], age],
      [FORM_XV_AP_WORKMAN_LABELS[2], marks],
      [FORM_XV_AP_WORKMAN_LABELS[3], father]
    ],
    tableHeaders: FORM_XV_AP_TABLE_HEADERS,
    tableRow: ['1', from, to, natureOfWork, rate, remarks]
  };
};

const strokeRect = (doc, x, y, w, h) => {
  doc.setDrawColor(0, 0, 0);
  doc.setLineWidth(0.7);
  doc.rect(x, y, w, h, 'S');
};

const drawFormXVAPCertificatePage = (doc, model) => {
  const pageWidth = doc.internal.pageSize.getWidth();
  const marginX = 28;
  const usableWidth = pageWidth - marginX * 2;
  const padX = 6;
  let y = 24;

  (model.titles || []).forEach((title, idx) => {
    const isFormName = /^form\s+/i.test(title);
    const isService = /service\s+certificate/i.test(title);
    const size = isFormName || isService ? 11 : 8;
    const bold = isFormName || isService || idx === 0;
    doc.setFont('helvetica', bold ? 'bold' : 'normal');
    doc.setFontSize(size);
    const lines = doc.splitTextToSize(String(title || ''), usableWidth - padX * 2);
    const h = Math.max(isFormName || isService ? 20 : 16, lines.length * (size + 2) + 8);
    strokeRect(doc, marginX, y, usableWidth, h);
    const textY = y + (h - lines.length * (size + 2)) / 2 + size;
    doc.text(lines, pageWidth / 2, textY, { align: 'center' });
    y += h;
  });

  const half = usableWidth / 2;
  (model.adminRows || []).forEach((pair, idx) => {
    const labels = model.adminLabels?.[idx] || ['', ''];
    const leftText = `${labels[0]}\n${pair[0] || ''}`;
    const rightText = `${labels[1]}\n${pair[1] || ''}`;
    doc.setFont('helvetica', 'normal');
    doc.setFontSize(9);
    const leftLines = doc.splitTextToSize(leftText, half - padX * 2);
    const rightLines = doc.splitTextToSize(rightText, half - padX * 2);
    const h = Math.max(52, Math.max(leftLines.length, rightLines.length) * 11 + 14);
    strokeRect(doc, marginX, y, usableWidth, h);
    doc.line(marginX + half, y, marginX + half, y + h);
    doc.text(leftLines, marginX + padX, y + 12);
    doc.text(rightLines, marginX + half + padX, y + 12);
    y += h;
  });

  (model.fieldRows || []).forEach(([label, value]) => {
    const labelW = Math.min(220, usableWidth * 0.38);
    doc.setFont('helvetica', 'normal');
    doc.setFontSize(9);
    const labelLines = doc.splitTextToSize(String(label || ''), labelW - padX * 2);
    const valueLines = doc.splitTextToSize(String(value || ''), usableWidth - labelW - padX * 2);
    const h = Math.max(22, Math.max(labelLines.length, valueLines.length) * 11 + 10);
    strokeRect(doc, marginX, y, usableWidth, h);
    doc.line(marginX + labelW, y, marginX + labelW, y + h);
    doc.text(labelLines, marginX + padX, y + 14);
    doc.text(valueLines, marginX + labelW + padX, y + 14);
    y += h;
  });

  y += 8;
  const headers = model.tableHeaders || FORM_XV_AP_TABLE_HEADERS;
  const dataRow = model.tableRow || [];
  const weights = [0.08, 0.14, 0.14, 0.24, 0.26, 0.14];
  const colXs = [];
  let x = marginX;
  weights.forEach((w) => {
    colXs.push({ x, w: w * usableWidth });
    x += w * usableWidth;
  });

  const drawTableRow = (cells, { header = false, minH = 22 } = {}) => {
    doc.setFont('helvetica', header ? 'bold' : 'normal');
    doc.setFontSize(8);
    const wrapped = cells.map((cell, i) =>
      doc.splitTextToSize(String(cell || ''), Math.max(18, colXs[i].w - 8))
    );
    const h = Math.max(minH, Math.max(...wrapped.map((lines) => lines.length)) * 10 + 10);
    wrapped.forEach((lines, i) => {
      strokeRect(doc, colXs[i].x, y, colXs[i].w, h);
      doc.text(lines, colXs[i].x + 4, y + 12);
    });
    y += h;
  };

  drawTableRow(headers, { header: true, minH: 28 });
  drawTableRow(dataRow, { minH: 26 });
};

/**
 * Fast Form XV Service Certificate PDF — one page per corresponding employee.
 * Draws from mapped employee rows (no Excel ZIP round-trip).
 */
export async function buildFormXVAPServiceCertificatePdfBlob({
  headerFormData = {},
  parsedFormHeader = {},
  employeeRows = [],
  fileName = '',
  title = ''
} = {}) {
  void title;
  const rows = filterFormXVAPCorrespondingEmployeeRows(employeeRows);
  if (!rows.length) {
    throw new Error('No corresponding employees found for Form XV Service Certificate PDF.');
  }

  const doc = new jsPDF({
    unit: 'pt',
    format: 'a4',
    orientation: 'portrait'
  });

  for (let i = 0; i < rows.length; i += 1) {
    if (i > 0) doc.addPage();
    const model = buildFormXVAPCertificatePageModel({
      headerFormData,
      parsedFormHeader,
      row: rows[i],
      fileName
    });
    drawFormXVAPCertificatePage(doc, model);
    if (i > 0 && i % 25 === 0) {
      await new Promise((resolve) => setTimeout(resolve, 0));
    }
  }

  return doc.output('blob');
}
