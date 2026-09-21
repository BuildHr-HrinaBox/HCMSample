import * as XLSX from 'xlsx';

const SMARTBROWZ_CONVERT_URL = '/server/sdkpdf_function/convert';

const escapeHtml = (value) =>
  String(value ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');

const extractTableHtml = (html) => {
  const raw = String(html || '');
  const match = raw.match(/<table[\s\S]*<\/table>/i);
  return match ? match[0] : raw;
};

const looksWideStatutoryName = (fileName = '', title = '') => {
  const blob = `${fileName} ${title}`.toLowerCase();
  return /form[\s._-]*t\b|form t|attendance|muster|wage register|form[\s._-]*[wpq]\b|form[\s._-]*xviii|form[\s._-]*xxvi/.test(
    blob
  );
};

export function inferSmartbrowzPdfOptions({ fileName = '', title = '', maxCols = 0 } = {}) {
  const wideName = looksWideStatutoryName(fileName, title);
  const veryWide = wideName || maxCols > 18;
  const wide = veryWide || maxCols > 10;
  const format = veryWide ? 'A2' : wide ? 'A3' : 'A4';
  return {
    pdf_options: {
      format,
      landscape: wide,
      print_background: true,
      display_header_footer: false,
      margin: {
        top: '8mm',
        bottom: '8mm',
        left: '6mm',
        right: '6mm'
      }
    },
    page_options: {
      javascript_enabled: false
    }
  };
}

export function workbookArrayBufferToHtml(arrayBuffer, { sheetTitle = '', fileLabel = '' } = {}) {
  const wb = XLSX.read(arrayBuffer, { type: 'array', cellDates: true, cellText: true });
  const names = wb.SheetNames || [];
  let maxCols = 0;
  const sections = names
    .map((name) => {
      const ws = wb.Sheets[name];
      if (!ws) return '';
      if (ws['!ref']) {
        try {
          const range = XLSX.utils.decode_range(ws['!ref']);
          maxCols = Math.max(maxCols, range.e.c - range.s.c + 1);
        } catch (_) {
          /* ignore */
        }
      }
      const table = extractTableHtml(XLSX.utils.sheet_to_html(ws, { header: '', footer: '' }));
      if (!table) return '';
      const heading = escapeHtml(fileLabel && names.length === 1 ? fileLabel : name);
      return `<section class="sheet"><h2>${heading}</h2>${table}</section>`;
    })
    .filter(Boolean);
  return { html: sections.join('\n'), maxCols, sheetTitle };
}

export function wrapStatutoryPdfHtml({ title = '', bodyHtml = '' } = {}) {
  const heading = String(title || '').trim();
  const titleHtml =
    heading && !/\.(xlsx|xls|xlsm|zip)$/i.test(heading)
      ? `<h1 class="form-title">${escapeHtml(heading)}</h1>`
      : '';
  return `<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8" />
  <style>
    * { box-sizing: border-box; }
    body {
      font-family: Arial, Helvetica, sans-serif;
      color: #111;
      margin: 0;
      padding: 8px;
    }
    h1.form-title {
      text-align: center;
      font-size: 14px;
      margin: 0 0 10px;
    }
    h2 {
      font-size: 11px;
      margin: 0 0 6px;
    }
    .sheet { page-break-after: always; }
    .sheet:last-child { page-break-after: auto; }
    table {
      border-collapse: collapse;
      width: auto;
      max-width: 100%;
      font-size: 8px;
    }
    td, th {
      border: 1px solid #222;
      padding: 2px 4px;
      vertical-align: top;
      white-space: nowrap;
    }
  </style>
</head>
<body>
  ${titleHtml}
  ${bodyHtml}
</body>
</html>`;
}

const blobLooksLikePdf = async (blob) => {
  if (!blob || blob.size < 80) return false;
  const type = String(blob.type || '').toLowerCase();
  if (type.includes('application/pdf')) return true;
  const header = new Uint8Array(await blob.slice(0, 5).arrayBuffer());
  return header[0] === 0x25 && header[1] === 0x50 && header[2] === 0x44 && header[3] === 0x46;
};

export async function requestSmartbrowzPdf({
  html = '',
  url = '',
  templateId = '',
  templateData = undefined,
  fileName = 'document.pdf',
  pdf_options = {},
  page_options = {}
} = {}) {
  if (typeof fetch !== 'function') {
    throw new Error('fetch unavailable for SmartBrowz PDF');
  }
  const controller = typeof AbortController === 'function' ? new AbortController() : null;
  const timer = controller ? setTimeout(() => controller.abort(), 90000) : null;
  let resp;
  try {
    resp = await fetch(SMARTBROWZ_CONVERT_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        html,
        url,
        template_id: templateId || undefined,
        template_data: templateData,
        fileName,
        pdf_options,
        page_options
      }),
      signal: controller ? controller.signal : undefined
    });
  } finally {
    if (timer) clearTimeout(timer);
  }
  const contentType = resp.headers.get('content-type') || '';
  if (!resp.ok) {
    let message = `SmartBrowz PDF failed (${resp.status})`;
    try {
      if (contentType.includes('json')) {
        const json = await resp.json();
        message = json.message || json.error || message;
      } else {
        const text = await resp.text();
        if (text) message = text.slice(0, 400);
      }
    } catch (_) {
      /* keep default message */
    }
    throw new Error(message);
  }
  const blob = await resp.blob();
  if (!(await blobLooksLikePdf(blob))) {
    throw new Error('SmartBrowz returned an empty or invalid PDF');
  }
  return blob;
}

export async function convertExcelFilesToSmartbrowzPdf({
  excelFiles = [],
  title = '',
  fileName = 'statutory-draft.pdf'
} = {}) {
  if (typeof process !== 'undefined' && process.env && process.env.NODE_ENV === 'test') {
    throw new Error('SmartBrowz skipped in tests');
  }
  if (!Array.isArray(excelFiles) || excelFiles.length === 0) {
    throw new Error('No Excel data for SmartBrowz PDF');
  }
  let maxCols = 0;
  const bodyParts = [];
  for (const file of excelFiles) {
    const { html, maxCols: cols } = workbookArrayBufferToHtml(file.arrayBuffer, {
      fileLabel: file.label || ''
    });
    maxCols = Math.max(maxCols, cols || 0);
    if (html) bodyParts.push(html);
  }
  if (!bodyParts.length) {
    throw new Error('No worksheet HTML for SmartBrowz PDF');
  }
  const html = wrapStatutoryPdfHtml({ title, bodyHtml: bodyParts.join('\n') });
  const inferred = inferSmartbrowzPdfOptions({ fileName, title, maxCols });
  return requestSmartbrowzPdf({
    html,
    fileName,
    pdf_options: inferred.pdf_options,
    page_options: inferred.page_options
  });
}
