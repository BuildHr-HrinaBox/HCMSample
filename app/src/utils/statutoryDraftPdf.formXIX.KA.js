/**
 * Karnataka Form XIX (Wage Slip) PDF normalization helpers.
 * Keeps Karnataka-only layout handling isolated from the shared Form XIX logic.
 */

const FORM_XIX_KA_WAGE_HEADERS = [
    'No. of days worked',
    'Rate of daily wages/piece - rate',
    'No. of units worked in case of piece rate',
    'Dates on which overtime worked',
    'Overtime hours and amount of overtime wages',
  ];
  
  const FORM_XIX_KA_FOOTER_LABELS = [
    'Gross wages payable',
    'Deductions, any',
    'If Actual wages paid',
  ];
  const FORM_XIX_KA_SIGNATURE_TEXT = 'Signature of the contractor or his Representative';
  
  const normalizeText = (text) =>
    String(text ?? '')
      .replace(/\s+/g, ' ')
      .trim();
  
  const normalizeLower = (text) => normalizeText(text).toLowerCase();
  
  const rowToText = (row) =>
    (Array.isArray(row) ? row : [])
      .map((cell) => normalizeText(cell))
      .filter(Boolean)
      .join(' ');
  
  const rowHasText = (row) => rowToText(row) !== '';
  
  const isHeaderRowBlob = (blob) => {
    const text = normalizeLower(blob);
    return (
      /no\.?\s*of\s+days\s+worked/.test(text) &&
      /rate\s+of\s+daily\s+wages|piece/.test(text) &&
      /overtime/.test(text)
    );
  };
  
  const isFooterLabelBlob = (blob) => {
    const text = normalizeLower(blob);
    return (
      /gross\s+wages\s+payable/.test(text) &&
      (/deductions?\s*,?\s*any/.test(text) || /actual\s+wages\s+paid/.test(text))
    );
  };
  const isGrossLabel = (text) => /gross\s+wages\s+payable/i.test(String(text || ''));
  const isDeductionsLabel = (text) => /deductions?\s*,?\s*any/i.test(String(text || ''));
  const isActualWagesLabel = (text) => /if\s+actual\s+wages\s+paid|actual\s+wages\s+paid/i.test(String(text || ''));
  
  const isSignatureBlob = (blob) =>
    /signature\s+of\s+(?:the\s+)?contractor(?:\s+or\s+his)?\s+representative/i.test(
      String(blob || '')
    );
  
  const isSexLine = (line) => /sex\s+and\s+identification\s+marks/i.test(line);
  const isTokenLine = (line) => /token\/?\s*ticket\s+no/i.test(line);
  const isPrincipalLine = (line) => /name\s+and\s+address\s+of\s+principal\s+employer/i.test(line);
  const isContractorLine = (line) => /name\s+and\s+address\s+of\s+contractor/i.test(line);
  const isEstablishmentLine = (line) =>
    /name\s+and\s+address\s+of\s+establishment|establishment\s+in\/under\s+which\s+contract\s+is\s+carried/i.test(
      line
    );
  const isNatureLine = (line) => /nature\s+of\s+work.*location\s+of\s+work|nature\s+and\s+location\s+of\s+work/i.test(line);
  const isRuleLine = (line) => /\[?\s*see\s+rule\s*78\s*\(1\)\(b\)\s*\]?/i.test(line);
  const isWageSlipLine = (line) => /^wages?\s+slip\b/i.test(line);
  const isFormTitleLine = (line) => /^form\s*xix\b/i.test(line);
  
  const toFiveCols = (values = []) => {
    const out = ['', '', '', '', ''];
    for (let i = 0; i < Math.min(5, values.length); i += 1) out[i] = String(values[i] ?? '').trim();
    return out;
  };
  
  const findMainHeaderRow = (rows, tableStartRow = 0) => {
    const start = Math.max(0, tableStartRow - 2);
    const end = Math.min(rows.length - 1, tableStartRow + 10);
    for (let r = start; r <= end; r += 1) {
      if (isHeaderRowBlob(rowToText(rows[r]))) return r;
    }
    for (let r = 0; r < Math.min(rows.length, 28); r += 1) {
      if (isHeaderRowBlob(rowToText(rows[r]))) return r;
    }
    return -1;
  };
  
  const findWageColumnIndexes = (headerRow = []) => {
    const cols = [];
    for (let c = 0; c < headerRow.length; c += 1) {
      const t = normalizeText(headerRow[c]);
      if (!t) continue;
      cols.push(c);
    }
    while (cols.length < 5) {
      cols.push(cols.length ? cols[cols.length - 1] + 1 : cols.length);
    }
    return cols.slice(0, 5);
  };
  
  const readAt = (rows, r, c) => normalizeText(rows?.[r]?.[c]);
  
  const FORM_XIX_KA_FIELD_SPECS = [
    {
      key: 'contractor',
      label: 'Name and address of contractor',
      test: isContractorLine,
    },
    {
      key: 'establishment',
      label: 'Name and address of establishment in/under which contract is carried on',
      test: isEstablishmentLine,
    },
    {
      key: 'nature',
      label: 'Nature of work and location of work',
      test: isNatureLine,
    },
    {
      key: 'principal',
      label: 'Name and address of Principal Employer, for the week/fortnight/month',
      test: isPrincipalLine,
    },
    {
      key: 'sex',
      label: 'Sex and identification marks',
      test: isSexLine,
    },
    {
      key: 'token',
      label: 'Token/Ticket No.',
      test: isTokenLine,
    },
  ];
  
  const collapsePunctuation = (text) =>
    String(text || '')
      .replace(/[.\u2026]{2,}/g, '.')
      .replace(/\s+/g, ' ')
      .trim();
  
  const isDotsOnlyValue = (text) => /^[.\u2026\s_-]+$/.test(String(text || '').trim());
  
  const splitLabelAndValue = (line) => {
    const src = normalizeText(line);
    if (!src) return { label: '', value: '' };
    const match = src.match(/^(.+?)(?:\s*[:\-]\s*|\s+[.\u2026]{3,}\s*|\s{2,})(.+)$/);
    if (!match) return { label: src, value: '' };
    const value = normalizeText(match[2]);
    return {
      label: collapsePunctuation(match[1]),
      value: isDotsOnlyValue(value) ? '' : value,
    };
  };
  
  const stripKaLabelPrefix = (line, label) => {
    const src = normalizeText(line);
    const prefix = collapsePunctuation(label).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    const stripped = src
      .replace(new RegExp(`^${prefix}\\s*[.:-]*\\s*`, 'i'), '')
      .replace(/^[.\u2026:\s-]+/, '')
      .trim();
    if (!stripped || isDotsOnlyValue(stripped) || stripped.toLowerCase() === src.toLowerCase()) {
      return '';
    }
    return stripped;
  };
  
  const formatKaFieldLine = (label, value) => {
    const l = collapsePunctuation(label).replace(/[.:]+$/g, '').trim();
    const v = normalizeText(value);
    if (!l && !v) return '';
    if (!v) return collapsePunctuation(label);
    if (/principal\s+employer/i.test(l)) return `${l} ${v}`.trim();
    if (/token\/?\s*ticket\s+no/i.test(l)) return `Token/Ticket No. ${v}`.trim();
    return `${l}. ${v}`.trim();
  };
  
  const isReservedKaMetaLine = (line) =>
    isFormTitleLine(line) ||
    isRuleLine(line) ||
    isWageSlipLine(line) ||
    isContractorLine(line) ||
    isEstablishmentLine(line) ||
    isNatureLine(line) ||
    isPrincipalLine(line) ||
    isSexLine(line) ||
    isTokenLine(line) ||
    isHeaderRowBlob(line) ||
    isFooterLabelBlob(line) ||
    isSignatureBlob(line);
  
  const collectPreHeaderLines = (rows, headerRow, fallbackMetaLines = []) =>
    [
      ...Array.from({ length: Math.max(0, headerRow) }, (_, r) => rowToText(rows[r])),
      ...(Array.isArray(fallbackMetaLines) ? fallbackMetaLines : []),
    ]
      .map((line) => normalizeText(line))
      .filter(Boolean);
  
  const resolveFieldValue = (allLines, predicate, canonicalLabel) => {
    for (let i = 0; i < allLines.length; i += 1) {
      const line = allLines[i];
      if (!predicate(line)) continue;
      const inline = stripKaLabelPrefix(line, canonicalLabel) || splitLabelAndValue(line).value;
      if (inline && !isReservedKaMetaLine(inline)) return inline;
      const next = normalizeText(allLines[i + 1] || '');
      if (next && !isReservedKaMetaLine(next) && !isDotsOnlyValue(next)) return next;
    }
    return '';
  };
  
  const isLikelyContractorValue = (line) => {
    const t = normalizeText(line);
    if (!t || isReservedKaMetaLine(t) || t.length > 90) return false;
    return /limited|pvt\.?|private|llc|llp|\bltd\.?\b|energy|company/i.test(t);
  };
  
  const isLikelyNatureValue = (line) => {
    const t = normalizeText(line);
    if (!t || isReservedKaMetaLine(t) || t.length > 80) return false;
    return /^(?:ka|karnataka)[-–\s]/i.test(t);
  };
  
  const normalizeKaNatureValue = (value) =>
    normalizeText(value).replace(/^(ka|karnataka)[-–]/i, (_, code) => `${String(code).toUpperCase()} `);
  
  const pickMetaLinesForKarnataka = (rows, headerRow, fallbackMetaLines = []) => {
    const allLines = collectPreHeaderLines(rows, headerRow, fallbackMetaLines);
    const unlabeled = allLines.filter((line) => !isReservedKaMetaLine(line) && !isDotsOnlyValue(line));
  
    const formTitle = allLines.find(isFormTitleLine) || 'FORM XIX';
    const ruleText = allLines.find(isRuleLine) || '[See rule 78 (1)(b)]';
    const wageSlip = allLines.find(isWageSlipLine) || 'Wages Slip';
  
    const values = {};
    FORM_XIX_KA_FIELD_SPECS.forEach((spec) => {
      values[spec.key] = resolveFieldValue(allLines, spec.test, spec.label);
    });
    if (!values.contractor) {
      const company = unlabeled.find(isLikelyContractorValue);
      if (company) values.contractor = company;
    }
    if (!values.nature) {
      const location = unlabeled.find(isLikelyNatureValue);
      if (location) values.nature = location;
    }
    if (values.nature) values.nature = normalizeKaNatureValue(values.nature);
  
    return [
      /^form\s*xix\b/i.test(formTitle) ? 'FORM XIX' : formTitle,
      /see\s+rule\s*78/i.test(ruleText) ? '[See rule 78 (1)(b)]' : ruleText,
      /^wages?\s+slip\b/i.test(wageSlip) ? 'Wages Slip' : wageSlip,
      ...FORM_XIX_KA_FIELD_SPECS.map((spec) => formatKaFieldLine(spec.label, values[spec.key])),
    ].filter(Boolean);
  };
  
  export const looksLikeFormXIXKarnatakaPdfContext = (
    metaLines = [],
    rows = [],
    sheetName = '',
    fileName = ''
  ) => {
    const blob = [...(metaLines || []), ...(rows || []).slice(0, 24).flat(), sheetName || '', fileName || '']
      .map((s) => normalizeLower(s))
      .join(' ');
    if (!blob) return false;
    const isXix = /form[\s._-]*xix\b/.test(blob) || (/wage\s*s?lip/.test(blob) && /rule\s*78/.test(blob));
    if (!isXix) return false;
    const kaName =
      /karnataka/.test(blob) ||
      /form[\s._-]*xix[\s._-]*(?:ka|karnataka)\b/.test(blob) ||
      /\bxix[_\s-]*ka\b/.test(blob) ||
      /\(ka\)/.test(blob);
    const otherStateForm =
      /form[\s._-]*xix[\s._-]*(?:gj|mp|ap|rj|tn)\b/.test(blob) ||
      /gujarat|madhya\s+pradesh|andhra\s+pradesh|rajasthan|tamil[\s._-]*nadu/.test(blob);
    if (otherStateForm && !kaName) return false;
    const hasKaFields =
      /sex\s+and\s+identification\s+marks/.test(blob) || /token\/?\s*ticket\s+no/.test(blob);
    if (!kaName && !hasKaFields) return false;
    return (
      /no\.?\s*of\s+days\s+worked/.test(blob) &&
      /overtime\s+hours|dates\s+on\s+which\s+overtime/.test(blob)
    );
  };
  
  export const normalizeFormXIXKarnatakaWageSlipPdfMatrix = (
    rows,
    colCount,
    tableStartRow = 0,
    metaLines = []
  ) => {
    const src = Array.isArray(rows) ? rows : [];
    if (!src.length) return { rows: src, colCount, tableStartRow, metaLines };
  
    const headerRow = findMainHeaderRow(src, tableStartRow);
    if (headerRow < 0) return { rows: src, colCount, tableStartRow, metaLines };
  
    const wageCols = findWageColumnIndexes(src[headerRow] || []);
    const headers = FORM_XIX_KA_WAGE_HEADERS.map((fallback, i) => readAt(src, headerRow, wageCols[i]) || fallback);
  
    let numberRow = -1;
    for (let r = headerRow + 1; r <= Math.min(src.length - 1, headerRow + 4); r += 1) {
      const cells = wageCols.map((c) => readAt(src, r, c)).filter(Boolean);
      if (cells.length >= 4 && cells.every((v) => /^\d{1,2}$/.test(v))) {
        numberRow = r;
        break;
      }
    }
  
    let footerLabelRow = -1;
    for (let r = Math.max(headerRow + 2, numberRow + 1); r < Math.min(src.length, headerRow + 16); r += 1) {
      if (isFooterLabelBlob(rowToText(src[r]))) {
        footerLabelRow = r;
        break;
      }
    }
  
    let footerValueRow = -1;
    if (footerLabelRow >= 0) {
      for (let r = footerLabelRow + 1; r <= Math.min(src.length - 1, footerLabelRow + 8); r += 1) {
        const text = rowToText(src[r]);
        if (!text || isSignatureBlob(text)) continue;
        if (/[.\u2026]{5,}/.test(text)) continue;
        // Prefer the first numeric-looking row under footer labels.
        if (/\b\d[\d,.-]*\b/.test(text)) {
          footerValueRow = r;
          break;
        }
        if (footerValueRow < 0) footerValueRow = r;
      }
    }
    if (footerValueRow < 0 && footerLabelRow >= 0) {
      for (let r = footerLabelRow + 1; r <= Math.min(src.length - 1, footerLabelRow + 8); r += 1) {
        const text = rowToText(src[r]);
        if (!text || isSignatureBlob(text) || /[.\u2026]{5,}/.test(text)) continue;
        footerValueRow = r;
        break;
      }
    }
  
    const dataRows = [];
    const dataStart = numberRow >= 0 ? numberRow + 1 : headerRow + 1;
    const dataEnd = footerLabelRow >= 0 ? footerLabelRow - 1 : Math.min(src.length - 1, dataStart + 8);
    for (let r = dataStart; r <= dataEnd; r += 1) {
      const row = src[r];
      if (!rowHasText(row)) continue;
      const blob = rowToText(row);
      if (isFooterLabelBlob(blob) || isSignatureBlob(blob)) continue;
      // Never allow header/meta/legal strings inside the data band.
      if (
        isRuleLine(blob) ||
        isWageSlipLine(blob) ||
        isContractorLine(blob) ||
        isEstablishmentLine(blob) ||
        isNatureLine(blob) ||
        isPrincipalLine(blob) ||
        isSexLine(blob) ||
        isTokenLine(blob)
      ) {
        continue;
      }
      const values = wageCols.map((c) => readAt(src, r, c));
      const valueBlob = normalizeLower(values.join(' '));
      if (
        isRuleLine(valueBlob) ||
        isWageSlipLine(valueBlob) ||
        /gross\s+wages|deductions|actual\s+wages|signature/.test(valueBlob)
      ) {
        continue;
      }
      if (values.every((v) => !v)) continue;
      dataRows.push(toFiveCols(values));
    }
  
    const footerLabelLine = ['', '', '', '', ''];
    const footerValueLine = ['', '', '', '', ''];
    const footerAnchorCols = [wageCols[0], wageCols[2], wageCols[4]];
    const footerLabelCols = { gross: -1, deductions: -1, actual: -1 };
    if (footerLabelRow >= 0) {
      const labelRow = Array.isArray(src[footerLabelRow]) ? src[footerLabelRow] : [];
      for (let c = 0; c < labelRow.length; c += 1) {
        const cell = readAt(src, footerLabelRow, c);
        if (!cell) continue;
        if (footerLabelCols.gross < 0 && isGrossLabel(cell)) footerLabelCols.gross = c;
        if (footerLabelCols.deductions < 0 && isDeductionsLabel(cell)) footerLabelCols.deductions = c;
        if (footerLabelCols.actual < 0 && isActualWagesLabel(cell)) footerLabelCols.actual = c;
      }
      FORM_XIX_KA_FOOTER_LABELS.forEach((fallback, i) => {
        const labelCol =
          i === 0
            ? footerLabelCols.gross
            : i === 1
              ? footerLabelCols.deductions
              : footerLabelCols.actual;
        const fromAnchor = readAt(src, footerLabelRow, labelCol >= 0 ? labelCol : footerAnchorCols[i]);
        footerLabelLine[i * 2] = fromAnchor || fallback;
      });
    } else {
      FORM_XIX_KA_FOOTER_LABELS.forEach((fallback, i) => {
        footerLabelLine[i * 2] = fallback;
      });
    }
    if (footerValueRow >= 0) {
      const colCandidates = [
        footerLabelCols.gross >= 0 ? footerLabelCols.gross : footerAnchorCols[0],
        footerLabelCols.deductions >= 0 ? footerLabelCols.deductions : footerAnchorCols[1],
        footerLabelCols.actual >= 0 ? footerLabelCols.actual : footerAnchorCols[2],
      ];
      const readValueUnderLabelCol = (startRow, col) => {
        if (col == null || col < 0) return '';
        for (let r = startRow; r <= Math.min(src.length - 1, startRow + 6); r += 1) {
          const line = rowToText(src[r]);
          if (!line || isSignatureBlob(line) || /[.\u2026]{5,}/.test(line)) continue;
          const v = readAt(src, r, col);
          if (!v) continue;
          if (isGrossLabel(v) || isDeductionsLabel(v) || isActualWagesLabel(v)) continue;
          return v;
        }
        return '';
      };
      const fromAnchor = colCandidates.map((c) => readValueUnderLabelCol(footerValueRow, c));
      if (fromAnchor.some(Boolean)) {
        footerValueLine[0] = fromAnchor[0] || '';
        footerValueLine[2] = fromAnchor[1] || '';
        footerValueLine[4] = fromAnchor[2] || '';
      } else {
        // Fallback for templates where only one packed numeric row exists.
        const packed = (src[footerValueRow] || []).map((c) => normalizeText(c)).filter(Boolean);
        if (packed.length >= 3) {
          footerValueLine[0] = packed[0] || '';
          footerValueLine[2] = packed[1] || '';
          footerValueLine[4] = packed[2] || '';
        } else if (packed.length === 2) {
          footerValueLine[0] = packed[0] || '';
          footerValueLine[2] = packed[1] || '';
        } else if (packed.length === 1) {
          // If a single value exists under footer, it is usually the deductions value.
          footerValueLine[2] = packed[0] || '';
        }
      }
    }
  
    let signatureText = '';
    const signatureScanStart = Math.max(footerValueRow, footerLabelRow, dataEnd, headerRow) + 1;
    for (let r = signatureScanStart; r < Math.min(src.length, signatureScanStart + 8); r += 1) {
      const line = rowToText(src[r]);
      if (!line) continue;
      if (isSignatureBlob(line)) {
        signatureText = line;
        break;
      }
    }
    if (!signatureText) signatureText = FORM_XIX_KA_SIGNATURE_TEXT;
  
    const compactRows = [
      headers,
      toFiveCols(['1', '2', '3', '4', '5']),
      ...(dataRows.length > 0 ? dataRows : [toFiveCols(['', 'Monthly Wages', '', 'NIL', 'NIL'])]),
    ];
  
    const reorderedMeta =
      pickMetaLinesForKarnataka(src, headerRow, metaLines).length > 0
        ? pickMetaLinesForKarnataka(src, headerRow, metaLines)
        : Array.isArray(metaLines)
          ? metaLines
          : [];
  
    return {
      rows: compactRows,
      colCount: 5,
      tableStartRow: 0,
      metaLines: reorderedMeta,
      formXIXKarnatakaLayout: true,
      formXIXKABottomSection: {
        labels: [
          footerLabelLine[0] || FORM_XIX_KA_FOOTER_LABELS[0],
          footerLabelLine[2] || FORM_XIX_KA_FOOTER_LABELS[1],
          footerLabelLine[4] || FORM_XIX_KA_FOOTER_LABELS[2],
        ],
        values: [footerValueLine[0] || '', footerValueLine[2] || '', footerValueLine[4] || ''],
        signatureText,
      },
    };
  };
  