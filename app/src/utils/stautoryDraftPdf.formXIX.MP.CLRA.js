const normalizeFormXIXMPText = (text) =>
    String(text || '')
      .replace(/\s+/g, ' ')
      .trim();
  
  const lowerFormXIXMPText = (text) => normalizeFormXIXMPText(text).toLowerCase();
  
  const FORM_XIX_MP_CLRA_HEADER_LABELS = [
    'Name and address of contractor',
    "Name and father's/husband's name of the workman",
    'Nature and location of work',
    'For the week/fortnight/month ending'
  ];
  
  const FORM_XIX_MP_CLRA_WAGE_LABELS = [
    'Number of days worked',
    'Number of units worked in case of piece rate workers',
    'Rate of daily wages/piece rate',
    'Amount of overtime wages',
    'Gross wages payable',
    'Deductions, If any',
    'Net amount of wages paid'
  ];
  
  const FORM_XIX_MP_CLRA_TITLES = ['FORM XIX', '[See rule 78 (1) (b)]', 'Wage Slip'];
  const FORM_XIX_MP_CLRA_FOOTER = 'Initial of the Contractor or his representative';
  
  const FORM_XIX_MP_HEADER_PATTERNS = [
    /name\s+and\s+address\s+of\s+(?:the\s+)?contractor/i,
    /name\s+and\s+(?:father.*husband|husband.*father).+workman/i,
    /nature\s+and\s+location\s+of\s+work/i,
    /for\s+the\s+(?:week|fortnight|month)/i
  ];
  
  const FORM_XIX_MP_WAGE_PATTERNS = [
    /(?:^|\b)(?:1[.)]?\s*)?(?:number|no\.?)\s+of\s+days\s+worked/i,
    /(?:^|\b)(?:2[.)]?\s*)?(?:number|no\.?)\s+of\s+units\s+worked(?:\s+in\s+case\s+of\s+piece\s+rate\s+workers?)?/i,
    /(?:^|\b)(?:3[.)]?\s*)?rate\s+of\s+daily\s+wages?\s*\/?\s*piece\s*-?\s*rate/i,
    /(?:^|\b)(?:4[.)]?\s*)?amount\s+of\s+overtime\s+wages/i,
    /(?:^|\b)(?:5[.)]?\s*)?gross\s+wages\s+payable/i,
    /(?:^|\b)(?:6[.)]?\s*)?deductions?\s*,?\s*if\s+any/i,
    /(?:^|\b)(?:7[.)]?\s*)?net\s+amount\s+of\s+wages\s+paid/i
  ];
  
  const compactFormXIXMPCandidates = (row) =>
    (Array.isArray(row) ? row : [])
      .map((cell) => normalizeFormXIXMPText(cell))
      .filter(Boolean);
  
  const looksLikeFormXIXTitle = (text) => /^form\s*xix\b/i.test(normalizeFormXIXMPText(text));
  const looksLikeRule78 = (text) => /see\s+rule\s*78/i.test(lowerFormXIXMPText(text));
  const looksLikeWageSlip = (text) => /^wage\s*s?lip$/i.test(normalizeFormXIXMPText(text));
  
  const findPatternValue = (rows = [], patterns = [], blockedPatterns = patterns) => {
    for (let r = 0; r < rows.length; r += 1) {
      const row = Array.isArray(rows[r]) ? rows[r] : [];
      for (let c = 0; c < row.length; c += 1) {
        const raw = normalizeFormXIXMPText(row[c]);
        if (!raw) continue;
  
        for (let p = 0; p < patterns.length; p += 1) {
          if (!patterns[p].test(raw)) continue;
  
          const inline = normalizeFormXIXMPText(raw.replace(/^[^:.-]*(?:[:.-]\s*|\s{2,})/, ''));
          if (inline && inline.toLowerCase() !== raw.toLowerCase()) {
            return { index: p, value: inline };
          }
  
          const rightCells = row
            .slice(c + 1)
            .map((cell) => normalizeFormXIXMPText(cell))
            .filter(Boolean);
          if (rightCells.length) {
            return { index: p, value: rightCells.join(' ') };
          }
  
          const nextRow = rows[r + 1] || [];
          const nextCells = compactFormXIXMPCandidates(nextRow);
          if (nextCells.length && !blockedPatterns.some((re) => re.test(nextCells[0]))) {
            return { index: p, value: nextCells.join(' ') };
          }
        }
      }
    }
    return null;
  };
  
  const collectMappedValues = (rows = [], patterns = [], blockedPatterns = patterns) => {
    const values = patterns.map(() => '');
    const workingRows = Array.isArray(rows) ? rows.map((row) => (Array.isArray(row) ? [...row] : row)) : [];
    for (let i = 0; i < patterns.length; i += 1) {
      const one = findPatternValue(workingRows, [patterns[i]], blockedPatterns);
      if (one && one.value) values[i] = one.value;
    }
    return values;
  };
  
  const sanitizeFormXIXMPWageValues = (values = []) => {
    const source = Array.isArray(values) ? values : [];
    return source.map((rawValue, currentIndex) => {
      let value = normalizeFormXIXMPText(rawValue);
      if (!value) return '';
  
      for (let wageIndex = 0; wageIndex < FORM_XIX_MP_WAGE_PATTERNS.length; wageIndex += 1) {
        if (wageIndex === currentIndex) continue;
        const pattern = FORM_XIX_MP_WAGE_PATTERNS[wageIndex];
        const match = value.match(pattern);
        if (!match) continue;
        const at = typeof match.index === 'number' ? match.index : -1;
        if (at <= 0) {
          value = '';
          break;
        }
        value = normalizeFormXIXMPText(value.slice(0, at));
      }
      return value;
    });
  };
  
  export const looksLikeFormXIXMPCLRAPdfContext = (metaLines = [], rows = [], sheetName = '') => {
    const blob = [...(metaLines || []), ...(rows || []).slice(0, 40).flat(), sheetName || '']
      .map(lowerFormXIXMPText)
      .join(' ');
    if (!blob) return false;
    const isXix = /form\s*xix\b|form[\s._-]*xix[\s._-]/.test(blob);
    if (!isXix) return false;
    const hasWageSlip = /wage\s*s?lip/.test(blob);
    const hasRule = /rule\s*78/.test(blob);
    const hasCoreRows =
      /name\s+and\s+address\s+of\s+contractor/.test(blob) &&
      /nature\s+and\s+location\s+of\s+work/.test(blob) &&
      /net\s+amount\s+of\s+wages/.test(blob);
    const hasMpHint =
      /madhya\s+pradesh|form[\s._-]*xix[\s._-]*mp|form_xix_mp|[-_\s]mp\b|\bmp[-_\s]|clra/.test(blob);
    const namedMpXix = /form[\s._-]*xix[\s._-]*mp|xix_mp/.test(blob);
    if (namedMpXix && (hasWageSlip || hasRule || hasCoreRows)) return true;
    return hasWageSlip && hasRule && hasCoreRows && hasMpHint;
  };
  
  export const normalizeFormXIXMPCLRAWageSlipPdfMatrix = (
    rows = [],
    colCount = 0,
    tableStartRow = 0,
    metaLines = []
  ) => {
    const sourceRows = Array.isArray(rows) ? rows : [];
    const sourceMeta = Array.isArray(metaLines) ? metaLines : [];
    const titleLines = [];
  
    sourceMeta.forEach((line) => {
      const text = normalizeFormXIXMPText(line);
      if (!text) return;
      if (looksLikeFormXIXTitle(text) || looksLikeRule78(text) || looksLikeWageSlip(text)) titleLines.push(text);
    });
  
    for (let r = 0; r < Math.min(sourceRows.length, Math.max(16, tableStartRow + 8)); r += 1) {
      const cells = compactFormXIXMPCandidates(sourceRows[r]);
      cells.forEach((text) => {
        if (looksLikeFormXIXTitle(text) || looksLikeRule78(text) || looksLikeWageSlip(text)) {
          titleLines.push(text);
        }
      });
    }
  
    const hasTitle = titleLines.some((line) => looksLikeFormXIXTitle(line));
    const hasRule = titleLines.some((line) => looksLikeRule78(line));
    const hasWage = titleLines.some((line) => looksLikeWageSlip(line));
    const orderedTitles = [];
    if (hasTitle) orderedTitles.push('FORM XIX');
    if (hasRule) orderedTitles.push('[See rule 78 (1) (b)]');
    if (hasWage) orderedTitles.push('Wage Slip');
    if (!orderedTitles.length) orderedTitles.push(...FORM_XIX_MP_CLRA_TITLES);
  
    const headerValues = collectMappedValues(sourceRows, FORM_XIX_MP_HEADER_PATTERNS, [
      ...FORM_XIX_MP_HEADER_PATTERNS,
      ...FORM_XIX_MP_WAGE_PATTERNS
    ]);
    const wageValues = sanitizeFormXIXMPWageValues(
      collectMappedValues(sourceRows, FORM_XIX_MP_WAGE_PATTERNS, FORM_XIX_MP_WAGE_PATTERNS)
    );
  
    // Titles stay in meta (above the box) — same pattern as Form XIV MP / Form XIX AP.
    const normalizedRows = [];

    FORM_XIX_MP_CLRA_HEADER_LABELS.forEach((label, idx) => {
      normalizedRows.push([label, headerValues[idx] || '']);
    });

    normalizedRows.push(['', '']);

    FORM_XIX_MP_CLRA_WAGE_LABELS.forEach((label, idx) => {
      normalizedRows.push([`${idx + 1} ${label}`, wageValues[idx] || '']);
    });

    return {
      rows: normalizedRows,
      colCount: 2,
      tableStartRow: 0,
      metaLines: orderedTitles,
      // Outside the bordered table (below), same as Form XIX AP signature/initials.
      formXIXAPFooterLines: [FORM_XIX_MP_CLRA_FOOTER],
      // Reuse Form XIX AP stacked draw path so titles render above the bordered table.
      formXIXAPLayout: true,
      formXIXMPCLRALayout: true,
      sourceColCount: Math.max(1, Number(colCount) || 1)
    };
  };
  
  
  