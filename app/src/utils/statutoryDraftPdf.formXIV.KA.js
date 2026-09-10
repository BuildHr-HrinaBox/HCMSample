/**
 * Karnataka Form XIV (Employment Card) specific normalization.
 * Kept separate from the shared Form XIV logic so other states are unaffected.
 */

const FORM_XIV_KA_HEADER_LABELS = [
    'Name and address if contractor',
    'Name and address of Establishment in/under which contract is carried on',
    'Nature and location of work',
    'Name and address of Principal Employer',
  ];
  
  const MAX_TRAILING_EMPTY_AFTER_CONTENT = 2;
  const FORM_XIV_KA_STACKED_VALUE_COL_INDEX = 4; // Excel column E
  const FORM_XIV_KA_ALT_VALUE_COL_INDEX = 3; // Excel column D (saved drafts)
  const FORM_XIV_KA_REMARKS_DOTS = ' ....................................................';
  const FORM_XIV_KA_SIGNATURE = 'Signature of Contractor';
  
  const normalizeText = (text) =>
    String(text || '')
      .replace(/\s+/g, ' ')
      .trim();
  
  const looksLikeFormXIVEmploymentCard = (metaLines, rows, sheetName = '') => {
    const blob = [...(metaLines || []), ...(rows || []).slice(0, 24).flat(), sheetName || '']
      .join(' ')
      .toLowerCase();
    if (!blob) return false;
    const isXiv =
      /form\s*xiv\b/.test(blob) ||
      /form[\s._-]*xiv[\s._-]/.test(blob) ||
      (/employment\s+card/.test(blob) && /see\s+rule\s*76/.test(blob));
    if (!isXiv) return false;
    return (
      /employment\s+card/.test(blob) ||
      /name\s+of\s+the\s+workman/.test(blob) ||
      /serial\s+(?:no\.?|number)\s+in\s+the\s+register/.test(blob)
    );
  };
  
  const rowText = (row) =>
    (Array.isArray(row) ? row : [])
      .map((c) => normalizeText(c))
      .filter(Boolean)
      .join(' ');
  
  const hasDottedFormXIVKarnatakaWorkmanLine = (rows = []) =>
    (Array.isArray(rows) ? rows : []).slice(0, 36).some((row) => {
      const t = rowText(row);
      return t && isWorkmanLabel(t) && /\.{4,}/.test(t);
    });
  
  export const looksLikeFormXIVKarnatakaPdfContext = (
    metaLines,
    rows,
    sheetName = '',
    fileName = ''
  ) => {
    const nameBlob = `${sheetName || ''} ${fileName || ''}`.trim();
    if (!looksLikeFormXIVEmploymentCard(metaLines, rows, nameBlob)) return false;
    const blob = [...(metaLines || []), ...(rows || []).slice(0, 24).flat(), nameBlob]
      .join(' ')
      .toLowerCase();
    const kaName =
      /karnataka/.test(blob) ||
      /form[\s._-]*xiv[\s._-]*ka\b/.test(blob) ||
      /\bxiv[_\s-]*ka\b/.test(blob) ||
      /\bform_xiv_ka\b/.test(blob) ||
      /form[\s._-]*xiv[\s._-]*karnataka/.test(blob);
    return kaName || hasDottedFormXIVKarnatakaWorkmanLine(rows);
  };
  
  const matchHeaderIndex = (text) => {
    const t = normalizeText(text).replace(/\.+$/g, '').replace(/\s*:?\s*$/, '').toLowerCase();
    if (!t) return -1;
    if (/name\s+and\s+address\s+(?:of|if)\s+(?:the\s+)?contractor/.test(t)) return 0;
    if (
      /name\s+and\s+address\s+of\s+(?:the\s+)?establishment/.test(t) ||
      /establishment\s+in\s*\/?\s*under\s+which/.test(t)
    ) {
      return 1;
    }
    if (/nature\s+and\s+location\s+of\s+work|nature\s+of\s+work\s+and\s+location/.test(t)) return 2;
    if (/name\s+and\s+address\s+of\s+(?:the\s+)?principal\s+employer|principal\s+employer/.test(t)) {
      return 3;
    }
    return -1;
  };
  
  const isWorkmanLabel = (text) => {
    const t = normalizeText(text).toLowerCase().replace(/^\d+[.)]?\s*/, '');
    if (!t) return false;
    return (
      /name\s+of\s+the\s+workman/.test(t) ||
      (/serial|s\.?\s*no/.test(t) && /register/.test(t)) ||
      (/nature/.test(t) && /employ|designat/.test(t)) ||
      (/wage/.test(t) && /rate|piece|particular|unit/.test(t)) ||
      /wage\s+period/.test(t) ||
      /tenure\s+of\s+employ|period\s+of\s+employ/.test(t) ||
      /^remarks?\b/.test(t) ||
      /date\s+of\s+entry\s+into\s+service/.test(t)
    );
  };
  
  const isHeaderLabel = (text) => matchHeaderIndex(text) >= 0;
  
  const extractInlineHeaderValue = (text, headerIndex) => {
    const src = normalizeText(text);
    if (!src) return '';
    const patterns = [
      /^name\s+and\s+address\s+(?:of|if)\s+(?:the\s+)?contractor(?:\s*[:.-]\s*|\s+\.{2,}\s*|\s{2,})?/i,
      /^name\s+and\s+address\s+of\s+(?:the\s+)?establishment\s+in\s*\/?\s*under\s+which\s+contract\s+is\s+carried\s+on(?:\s*[:.-]\s*|\s+\.{2,}\s*|\s{2,})?/i,
      /^nature\s+and\s+location\s+of\s+work(?:\s*[:.-]\s*|\s+\.{2,}\s*|\s{2,})?/i,
      /^name\s+and\s+address\s+of\s+(?:the\s+)?principal\s+employer(?:\s*[:.-]\s*|\s+\.{2,}\s*|\s{2,})?/i,
    ];
    const stripped = src.replace(patterns[headerIndex], '').replace(/^[.\-:\s]+/, '').trim();
    if (!stripped || stripped.toLowerCase() === src.toLowerCase()) return '';
    return stripped;
  };
  
  const splitInlineRemarksValue = (labelText) => {
    const src = normalizeText(labelText);
    if (!src) return { label: '', value: '' };
    const match = src.match(/^((?:\d{1,2}[.)]?\s*)?remarks?)(?:\s*[:.-]\s*|\s+\.{2,}\s*|\s{2,})(.+)$/i);
    if (!match) return { label: src, value: '' };
    return {
      label: normalizeText(match[1]),
      value: normalizeText(String(match[2] || '').replace(/^[.\-:\s]+/, '')),
    };
  };
  
  const splitInlineWorkmanValue = (text) => {
    const src = normalizeText(text);
    if (!src) return { label: '', value: '' };
    const wageRateFull = src.match(
      /^(\d{1,2}[.)]?\s*wage\s+rate\s*\([^)]*\))\s*(?:[:.-]\s*|\s+\.{2,}\s*|\s{2,})(.+)$/i
    );
    if (wageRateFull) {
      return {
        label: normalizeText(wageRateFull[1]).replace(/\.+$/g, ''),
        value: normalizeText(String(wageRateFull[2] || '').replace(/^[.\-:\s]+/, '')),
      };
    }
    const patterns = [
      /^(\d{1,2}[.)]?\s*name\s+of\s+the\s+workman)\b/i,
      /^(\d{1,2}[.)]?\s*serial\s+no\.?\s+in\s+the\s+register\s+of\s+workmen\s+employed)\b/i,
      /^(\d{1,2}[.)]?\s*nature\s+of\s+employment\s*\/\s*designation)\b/i,
      /^(\d{1,2}[.)]?\s*wage\s+rate(?:\s*\([^)]*\))?)\b/i,
      /^(\d{1,2}[.)]?\s*wage\s+period)\b/i,
      /^(\d{1,2}[.)]?\s*tenure\s+of\s+employment)\b/i,
      /^(\d{1,2}[.)]?\s*remarks?)\b/i,
    ];
    for (const re of patterns) {
      const m = src.match(re);
      if (!m) continue;
      const label = normalizeText(m[1]).replace(/\.+$/g, '');
      const value = normalizeText(src.slice(m[0].length).replace(/^[.\-:\s]+/, ''));
      if (!value || /^[_\s.-]{3,}$/.test(value)) return { label, value: '' };
      return { label, value };
    }
    return { label: src, value: '' };
  };
  
  const isSystemGeneratedDocumentNote = (text) =>
    /system\s+generated\s+document|computer\s+generated\s+document/i.test(String(text || ''));
  
  export const isFormXIVKarnatakaSignatureText = (text) =>
    /signature\s+of\s+(?:the\s+)?contractor/i.test(String(text || ''));
  
  const isFormXIVKaTitleLine = (text) => {
    const t = normalizeText(text);
    if (!t) return false;
    return (
      /^form\s*xiv\b/i.test(t) ||
      /^employment\s+card$/i.test(t) ||
      /see\s+rule\s*76/i.test(t)
    );
  };
  
  const assignHeaderValue = (headerValues, idx, value) => {
    if (idx < 0 || idx >= headerValues.length) return;
    const text = normalizeText(value);
    if (text && !headerValues[idx]) headerValues[idx] = text;
  };
  
  export const normalizeFormXIVKarnatakaPdfMatrix = (
    rows,
    colCount,
    tableStartRow = 0,
    incomingMetaLines = []
  ) => {
    const src = Array.isArray(rows) ? rows : [];
    if (!src.length) {
      return {
        rows: src,
        colCount,
        tableStartRow,
        metaLines: null,
        formXIVFooterLines: [FORM_XIV_KA_SIGNATURE],
        formXIVKALayout: true,
      };
    }
  
    const cellAt = (row, c) => normalizeText(row?.[c]);
    const metaLines = [];
    const headerValues = FORM_XIV_KA_HEADER_LABELS.map(() => '');
    const workmanRows = [];
    const trailingRows = [];
    const pendingOrphanValues = [];
    let sawRemarksRow = false;
  
    (Array.isArray(incomingMetaLines) ? incomingMetaLines : []).forEach((line) => {
      const text = normalizeText(line);
      if (!text) return;
      if (isFormXIVKaTitleLine(text)) {
        if (!metaLines.some((existing) => existing.toLowerCase() === text.toLowerCase())) {
          metaLines.push(text);
        }
        return;
      }
      const hdrIdx = matchHeaderIndex(text);
      if (hdrIdx >= 0) {
        assignHeaderValue(headerValues, hdrIdx, extractInlineHeaderValue(text, hdrIdx));
      }
    });
  
    const splitStackedRow = (row) => {
      const cells = Array.isArray(row) ? row : [];
      const nonEmpty = [];
      for (let c = 0; c < cells.length; c += 1) {
        const t = cellAt(cells, c);
        if (t) nonEmpty.push({ c, t });
      }
      if (!nonEmpty.length) return { left: '', value: '', kind: 'empty', headerIndex: -1 };
  
      let ordinal = '';
      let label = '';
      let labelCol = -1;
      let value = '';
      let kind = 'other';
      let headerIndex = -1;
  
      for (const { c, t } of nonEmpty) {
        if (/^\d{1,2}[.)]?$/.test(t)) {
          ordinal = t.replace(/\)/, '.');
          if (!/\.$/.test(ordinal)) ordinal = `${ordinal}.`;
          continue;
        }
        const combined = t.match(/^(\d{1,2})[.)]?\s+(.+)$/);
        if (combined && isWorkmanLabel(combined[2])) {
          ordinal = `${combined[1]}.`;
          label = combined[2].trim();
          labelCol = c;
          kind = 'workman';
          continue;
        }
        if (!label && isWorkmanLabel(t)) {
          label = t.replace(/\.+$/g, '').trim();
          labelCol = c;
          kind = 'workman';
          continue;
        }
        const hdrIdx = matchHeaderIndex(t);
        if (!label && hdrIdx >= 0) {
          label = FORM_XIV_KA_HEADER_LABELS[hdrIdx];
          labelCol = c;
          kind = 'header';
          headerIndex = hdrIdx;
          if (!value) {
            const inline = extractInlineHeaderValue(t, hdrIdx);
            if (inline) value = inline;
          }
        }
      }
  
      const stackedCandidates = [
        cellAt(cells, FORM_XIV_KA_STACKED_VALUE_COL_INDEX),
        cellAt(cells, FORM_XIV_KA_ALT_VALUE_COL_INDEX),
      ];
      for (const stacked of stackedCandidates) {
        if (
          stacked &&
          stacked !== label &&
          !isWorkmanLabel(stacked) &&
          !isHeaderLabel(stacked) &&
          !/^\d{1,2}[.)]?$/.test(stacked)
        ) {
          value = stacked;
          break;
        }
      }
  
      if (!value) {
        for (let i = nonEmpty.length - 1; i >= 0; i -= 1) {
          const { c, t } = nonEmpty[i];
          if (c === labelCol) continue;
          if (/^\d{1,2}[.)]?$/.test(t)) continue;
          if (t === label) continue;
          if (isWorkmanLabel(t) || isHeaderLabel(t)) continue;
          if (/^\.+$/.test(t) || /^[_\s.-]{3,}$/.test(t)) continue;
          value = t;
          break;
        }
      }
  
      if (!label && !value) {
        return { left: nonEmpty.map((x) => x.t).join(' '), value: '', kind: 'title', headerIndex: -1 };
      }
  
      let left = label || '';
      if (kind === 'workman' && ordinal && left && !/^\d{1,2}[.)]/.test(left)) {
        left = `${ordinal} ${left}`.replace(/\.\s+/, '. ');
      } else if (ordinal && !left) {
        left = ordinal;
      }
  
      if (kind === 'workman' && !value && /^(\d{1,2}[.)]\s*)?remarks?\b/i.test(left)) {
        const inlineRemarks = splitInlineRemarksValue(left);
        if (inlineRemarks.value) {
          left = inlineRemarks.label || left;
          value = inlineRemarks.value;
        }
      }
      if (kind === 'workman' && !value) {
        const inlineWorkman = splitInlineWorkmanValue(left);
        if (inlineWorkman.value) {
          left = inlineWorkman.label || left;
          value = inlineWorkman.value;
        }
      }
      if (kind === 'workman') {
        const isRemarks = /^(\d{1,2}[.)]\s*)?remarks?\b/i.test(left);
        if (isRemarks && !value) {
          const remarksBase = left.replace(/\s*\.{2,}.*$/, '').replace(/\.+$/g, '').trim();
          left = `${remarksBase}${FORM_XIV_KA_REMARKS_DOTS}`;
        } else {
          left = left.replace(/\s*\.{3,}.*$/, '').replace(/\.+$/g, '').trim();
        }
      }
  
      return { left, value, kind, headerIndex };
    };
  
    for (let r = 0; r < src.length; r += 1) {
      const row = src[r] || [];
      const filled = row.map((c) => normalizeText(c)).filter(Boolean);
      if (!filled.length) continue;
      const blob = filled.join(' ').toLowerCase();
  
      if (isSystemGeneratedDocumentNote(filled.join(' '))) {
        if (!sawRemarksRow) trailingRows.push([filled[0], '']);
        continue;
      }
      if (isFormXIVKarnatakaSignatureText(blob)) {
        continue;
      }
  
      const pair = splitStackedRow(row);
      if (pair.kind === 'header' || pair.headerIndex >= 0) {
        const idx = pair.headerIndex >= 0 ? pair.headerIndex : matchHeaderIndex(pair.left);
        if (idx >= 0 && pair.value && !headerValues[idx]) headerValues[idx] = pair.value;
        continue;
      }
  
      if (pair.kind === 'workman') {
        while (pendingOrphanValues.length) {
          const emptyIdx = headerValues.findIndex((v) => !String(v || '').trim());
          if (emptyIdx < 0) break;
          assignHeaderValue(headerValues, emptyIdx, pendingOrphanValues.shift());
        }
        workmanRows.push([pair.left, pair.value]);
        if (/^(\d{1,2}[.)]\s*)?remarks?\b/i.test(String(pair.left || '').trim())) sawRemarksRow = true;
        continue;
      }
  
      // Continuation value row (e.g. Serial No value on next line) should fill the
      // previous workman row's right-side value column.
      if (!sawRemarksRow && workmanRows.length > 0) {
        const last = workmanRows[workmanRows.length - 1];
        const joined = normalizeText(filled.join(' '));
        const lowerJoined = joined.toLowerCase();
        if (
          !String(last?.[1] || '').trim() &&
          joined &&
          !/^[_\s.-]{3,}$/.test(joined) &&
          !isWorkmanLabel(joined) &&
          !isHeaderLabel(joined) &&
          !/^form\s*xiv\b|^employment\s+card$|^see\s+rule/i.test(lowerJoined) &&
          !isFormXIVKarnatakaSignatureText(joined)
        ) {
          last[1] = joined.replace(/^[.\-:\s]+/, '').trim();
          continue;
        }
      }
  
      if (sawRemarksRow) continue;
  
      if (workmanRows.length === 0) {
        const titleLine = pair.left || filled.join(' ');
        if (matchHeaderIndex(titleLine) >= 0) continue;
        if (isFormXIVKaTitleLine(titleLine)) {
          if (!metaLines.some((existing) => existing.toLowerCase() === titleLine.toLowerCase())) {
            metaLines.push(titleLine);
          }
          continue;
        }
        if (titleLine && !isWorkmanLabel(titleLine) && !isHeaderLabel(titleLine)) {
          pendingOrphanValues.push(titleLine);
          continue;
        }
      }
  
      if (workmanRows.length === 0) {
        continue;
      }
      trailingRows.push([pair.left || filled.join(' '), pair.value || '']);
    }
  
    while (pendingOrphanValues.length) {
      const emptyIdx = headerValues.findIndex((v) => !String(v || '').trim());
      if (emptyIdx < 0) break;
      assignHeaderValue(headerValues, emptyIdx, pendingOrphanValues.shift());
    }
  
    const headerRows = FORM_XIV_KA_HEADER_LABELS.map((label, idx) => [
      `${label}:`,
      headerValues[idx] || '',
    ]);
  
    const compact = [...headerRows, ...workmanRows, ...trailingRows];
    let lastContent = -1;
    for (let r = 0; r < compact.length; r += 1) {
      if ((compact[r] || []).some((c) => normalizeText(c))) lastContent = r;
    }
    const trimmed =
      lastContent < 0
        ? compact
        : compact.slice(0, Math.min(compact.length, lastContent + 1 + MAX_TRAILING_EMPTY_AFTER_CONTENT));
  
    return {
      rows: trimmed.length ? trimmed : headerRows,
      colCount: 2,
      tableStartRow: 0,
      metaLines: metaLines.length ? metaLines : null,
      formXIVFooterLines: [FORM_XIV_KA_SIGNATURE],
      formXIVKALayout: true,
    };
  };