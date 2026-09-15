/**
 * Karnataka Form T — Combined Muster Roll cum Register of Wages.
 * Citation lines mention "Form XIII of Rules 75", which must never route this
 * sheet through the AP Form XIII Register of Workmen PDF layout.
 */

export const FORM_T_KA_PDF_TITLE = 'FORM T';
export const FORM_T_KA_PDF_SUBTITLE = 'COMBINED MUSTER ROLL CUM REGISTER OF WAGES';
export const FORM_T_KA_PDF_RULE =
  '[See Rule 24(9-B) of Karnataka Shops & Commercial Establishment Rules, 1963]';
export const FORM_T_KA_PDF_IN_LIEU_LINES = [
  'in lieu of',
  '1. Form I, II of Rule 22(4); Form IV of Rule 29(2); Forms V & VII of Rule 29(1) & (5) of Karnataka Minimum Wages Rules, 1958',
  '2. Form I of Rules 3(1) of Karnataka Payment of Wages Rules, 1963',
  '3. Form XIII of Rules 75; Form XV, XVII, XX, XXI, XXII, XXIII of 78(1)(a)(i), (ii) & (iii) of Karnataka Contract Labour (Regulation & Abolition) Rules, 1974',
  '4. Form XIII of Rule 43; Forms XVII, XVIII, XIX, XX, XXI, XXII of Rule 46(2)(a),(c) & (d) of Inter-state Migrant Workmen (Regulation of Employment and conditions of service) Karnataka Rules, 1981',
];

/** Official KA identity leaves (image-2 Excel model). */
export const FORM_T_KA_PDF_IDENTITY_HEADERS = [
  'S.NO',
  'Name and address of principal employer',
  "Father / Husband's Name",
  'Gender',
  'Designation / Department',
  'Date of Joining',
  'ESI No.',
  'UAN No.',
  'Wages fixed including VDA',
];

export const FORM_T_KA_PDF_ATTENDANCE_LABEL =
  'ATTENDANCE (Please mention the date of suspension of employees, if any)';

export const FORM_T_KA_PDF_EARNED_WAGES_GROUP = 'Earned wages and other allowances';
export const FORM_T_KA_PDF_DEDUCTIONS_GROUP = 'Deductions';

export const FORM_T_KA_PDF_PAYABLE_OT_HEADERS = [
  'No. of payable days',
  'Total OT hours',
];

export const FORM_T_KA_PDF_EARNED_WAGE_LEAVES = [
  'BASIC',
  'DA/VDA',
  'HRA',
  'Conveyance',
  'Medical Allowance',
  'Attendance Bonus',
  'Special Allowance',
  'OT',
  'NFH',
  'Maternity Benefit',
  'Others',
  'Subsistence Allowance if any',
  'Total',
];

export const FORM_T_KA_PDF_DEDUCTION_LEAVES = [
  'ESI',
  'PF',
  'PT',
  'Society',
  'Insurance',
  'Salary Advance',
  'Fines',
  'Damages',
  'Others',
  'Total',
];

export const FORM_T_KA_PDF_TRAILING_HEADERS = [
  'Net Amount Payable',
  'Mode of Payment Cash/ Cheque No.',
  'Employee signature or thumb impression',
];

export const FORM_T_KA_PDF_DATE_LABEL = 'Date:';
export const FORM_T_KA_PDF_SIGNATORY_LABEL = 'Authorised Signatory';

export const FORM_T_KA_PDF_WAGE_LEAVES = [
  ...FORM_T_KA_PDF_PAYABLE_OT_HEADERS,
  ...FORM_T_KA_PDF_EARNED_WAGE_LEAVES,
  ...FORM_T_KA_PDF_DEDUCTION_LEAVES,
  ...FORM_T_KA_PDF_TRAILING_HEADERS,
];

const normalizeText = (value) =>
  String(value || '')
    .replace(/\s+/g, ' ')
    .trim();

const blobFrom = (metaLines, rows, sheetName = '', fileName = '') =>
  [...(metaLines || []), ...(rows || []).slice(0, 18).flat(), sheetName || '', fileName || '']
    .map(normalizeText)
    .join(' ')
    .toLowerCase();

const isFormTFilename = (sheetName = '', fileName = '') => {
  const name = `${sheetName || ''} ${fileName || ''}`.toLowerCase();
  return (
    /form[\s._-]*t[\s._-]*ka\b/.test(name) ||
    /form[\s._-]*t[\s._-]*karnataka/.test(name) ||
    /form_{1,}t(?:_|-|$)/.test(name) ||
    /form[\s._-]+t[\s._-]+s[\s._-]*e\b/.test(name)
  );
};

export const looksLikeFormTSEKarnatakaPdfContext = (
  metaLines = [],
  rows = [],
  sheetName = '',
  fileName = ''
) => {
  if (isFormTFilename(sheetName, fileName)) return true;
  const blob = blobFrom(metaLines, rows, sheetName, fileName);
  if (!blob) return false;
  const isFormT = /(?:^|[^a-z0-9])form[\s._-]*t(?:[^a-z0-9]|$)/.test(blob);
  const isCombinedMuster = /combined\s+muster\s+roll\s+cum\s+register\s+of\s+wages/.test(blob);
  const isRule249B = /rule\s+24\s*\(\s*9[\s-]*b/.test(blob);
  if (isCombinedMuster && (isFormT || /karnataka/.test(blob) || isRule249B)) return true;
  if (isFormT && (isCombinedMuster || isRule249B)) return true;
  return false;
};

export const isFormTSEKarnatakaTableHeaderRow = (row = []) => {
  const blob = (Array.isArray(row) ? row : []).map(normalizeText).join(' ').toLowerCase();
  if (!blob) return false;
  const hasSerial = /(?:^|[^a-z])s\.?\s*no(?:[^a-z]|$)/.test(blob);
  const hasIdentity =
    /name of(?:\s+the)?\s+employee/.test(blob) ||
    /father/.test(blob) ||
    /principal\s+employer/.test(blob) ||
    /designation/.test(blob) ||
    /attendance/.test(blob);
  return hasSerial && hasIdentity;
};

export const isFormTSEKarnatakaAdminRowBlob = (blob) => {
  const t = String(blob || '').toLowerCase();
  if (!t) return false;
  if (isFormTSEKarnatakaTableHeaderRow([t])) return false;
  return (
    /month\s*\/\s*year/.test(t) ||
    /name\s+and\s+address\s+of\s+the\s+establishment/.test(t) ||
    /name\s+and\s+address\s+of\s+(?:the\s+)?employer/.test(t) ||
    (/address\s+of\s+the\s+establishment/.test(t) && !/principal/.test(t)) ||
    (/name\s+and\s+address\s+of\s+(?:the\s+)?contractor/.test(t) && !/principal/.test(t)) ||
    /nature\s+of\s+work\s+and\s+location/.test(t)
  );
};

const collectLines = (metaLines, rows, tableStart) => {
  const fromRows = (rows || [])
    .slice(0, Math.min(Math.max(Number(tableStart) || 0, 12), (rows || []).length))
    .map((row) =>
      (Array.isArray(row) ? row : [])
        .map(normalizeText)
        .filter(Boolean)
        .join(' ')
    )
    .filter(Boolean);
  const fromMeta = (metaLines || []).map(normalizeText).filter(Boolean);
  const seen = new Set();
  const out = [];
  [...fromMeta, ...fromRows].forEach((line) => {
    const key = line.toLowerCase();
    if (!key || seen.has(key)) return;
    seen.add(key);
    out.push(line);
  });
  return out;
};

const pickLine = (lines, pattern, fallback = '') => {
  const found = (lines || []).find((line) => pattern.test(line));
  return found ? normalizeText(found) : fallback;
};

export const getFormTSEKarnatakaHeaderTitles = (metaLines = [], rows = [], tableStart = 0) => {
  const lines = collectLines(metaLines, rows, tableStart);
  const formTitle = pickLine(lines, /^form\s*t\b/i, FORM_T_KA_PDF_TITLE);
  const subtitle = pickLine(
    lines,
    /combined\s+muster\s+roll|muster\s+roll\s+cum\s+register/i,
    FORM_T_KA_PDF_SUBTITLE
  );
  const rule = pickLine(lines, /rule\s+24\s*\(\s*9[\s-]*b/i, FORM_T_KA_PDF_RULE);
  const inLieu = pickLine(lines, /^in\s+lieu\s+of$/i, FORM_T_KA_PDF_IN_LIEU_LINES[0]);
  const citation1 = pickLine(
    lines,
    /^\s*1\.\s*form\b/i,
    FORM_T_KA_PDF_IN_LIEU_LINES[1]
  );
  const citation2 = pickLine(
    lines,
    /^\s*2\.\s*form\b/i,
    FORM_T_KA_PDF_IN_LIEU_LINES[2]
  );
  const citation3 = pickLine(
    lines,
    /^\s*3\.\s*form\b/i,
    FORM_T_KA_PDF_IN_LIEU_LINES[3]
  );
  const citation4 = pickLine(
    lines,
    /^\s*4\.\s*form\b/i,
    FORM_T_KA_PDF_IN_LIEU_LINES[4]
  );
  return [
    /^form\s*t\b/i.test(formTitle) ? FORM_T_KA_PDF_TITLE : formTitle,
    /combined\s+muster|muster\s+roll\s+cum\s+register/i.test(subtitle)
      ? FORM_T_KA_PDF_SUBTITLE
      : subtitle,
    rule,
    inLieu,
    citation1,
    citation2,
    citation3,
    citation4,
  ].filter(Boolean);
};

const FIELD_SPECS = [
  {
    key: 'monthYear',
    label: 'Month / Year',
    pattern: /month\s*\/\s*year/i,
  },
  {
    key: 'establishment',
    label: 'Name and address of the Establishment',
    pattern:
      /(?:name\s+and\s+)?address\s+of\s+the\s+establishment|name\s+and\s+address\s+of\s+(?:the\s+)?contractor/i,
  },
  {
    key: 'employer',
    label: 'Name and Address of employer',
    pattern:
      /name\s+and\s+address\s+of\s+(?:the\s+)?employer|nature\s+of\s+work\s+and\s+location\s+of\s+work/i,
  },
];

const splitFieldValue = (text, spec) => {
  const src = normalizeText(text);
  if (!src || !spec.pattern.test(src)) return '';
  const cut = src.replace(spec.pattern, '').replace(/^\s*[:.\-]+\s*/, '').trim();
  if (!cut || (spec.pattern.test(cut) && cut.length < 8)) return '';
  return cut;
};

const isTitleOrCitationLine = (text) => {
  const t = normalizeText(text);
  if (!t) return false;
  if (/^form\s*t\b/i.test(t)) return true;
  if (/combined\s+muster\s+roll|muster\s+roll\s+cum\s+register\s+of\s+wages/i.test(t)) return true;
  if (/rule\s+24\s*\(\s*9[\s-]*b/i.test(t)) return true;
  if (/^in\s+lieu\s+of$/i.test(t)) return true;
  if (/^\s*[1-4]\.\s*form\b/i.test(t)) return true;
  return false;
};

const isFormTHeaderNoiseLine = (text) => {
  const t = normalizeText(text);
  if (!t) return true;
  if (isTitleOrCitationLine(t)) return true;
  if (isFormTSEKarnatakaTableHeaderRow([t])) return true;
  if (/name\s+and\s+address\s+of\s+principal\s+employer/i.test(t) && /s\.?\s*no/i.test(t)) {
    return true;
  }
  return false;
};

export const extractFormTSEKarnatakaHeaderFields = (
  metaLines = [],
  rows = [],
  tableStart = 0
) => {
  const lines = collectLines(metaLines, rows, tableStart);
  const values = { monthYear: '', establishment: '', employer: '' };
  lines.forEach((line, idx) => {
    if (isFormTHeaderNoiseLine(line)) return;
    FIELD_SPECS.forEach((spec) => {
      if (values[spec.key]) return;
      if (!spec.pattern.test(line)) return;
      let inline = splitFieldValue(line, spec);
      if (!inline) {
        const next = lines[idx + 1];
        if (
          next &&
          !FIELD_SPECS.some((other) => other.pattern.test(next)) &&
          !isFormTHeaderNoiseLine(next)
        ) {
          inline = next;
        }
      }
      if (inline) values[spec.key] = inline;
    });
  });
  return FIELD_SPECS.map((spec) =>
    values[spec.key] ? `${spec.label} : ${values[spec.key]}` : `${spec.label} :`
  );
};

const isLeakedFormXiiiTitle = (text) => {
  const t = normalizeText(text).toLowerCase();
  if (!t) return false;
  if (/register\s+of\s+workmen\s+employed\s+by\s+contractor/.test(t) && !/in\s+lieu/.test(t)) {
    return true;
  }
  if (/vide\s+rule\s*75/.test(t) && /central\s*\/\s*a\.?p/i.test(t)) return true;
  return false;
};

const isLoneDateFooter = (row) => {
  const filled = (Array.isArray(row) ? row : []).map(normalizeText).filter(Boolean);
  return filled.length === 1 && /^date\s*:?\s*$/i.test(filled[0]);
};

const isFormTPdfDateFooterText = (text) => {
  const t = normalizeText(text);
  if (!t) return false;
  if (/date\s+of\s+joining|date\s+of\s+birth|date\s+of\s+suspension|date\s+of\s+payment/i.test(t)) {
    return false;
  }
  return /^date\s*:/i.test(t) || /^date\s*$/i.test(t);
};

const isFormTPdfSignatoryFooterText = (text) => {
  const t = normalizeText(text);
  if (!t) return false;
  if (/employee\s+signature|thumb\s+impression/i.test(t)) return false;
  return (
    /authori[sz]ed\s+signatory/i.test(t) ||
    /signature\s+of\s+(?:the\s+)?employer/i.test(t) ||
    /^for\s*\(/i.test(t)
  );
};

const isFormTPdfFooterCellText = (text) =>
  isFormTPdfDateFooterText(text) || isFormTPdfSignatoryFooterText(text);

/** Date: / Authorised Signatory rows belong under the table box, not inside it. */
export const isFormTSEKarnatakaPdfFooterOnlyRow = (row) => {
  if (!Array.isArray(row)) return false;
  if (
    looksLikePdfEmployeeName(cellAt(row, 1)) ||
    looksLikePdfEmployeeName(cellAt(row, FORM_T_KA_PDF_ATTENDANCE_START_COL0 + 1))
  ) {
    return false;
  }
  const filled = row.map(normalizeText).filter(Boolean);
  if (!filled.length) return false;
  if (isLoneDateFooter(row)) return true;
  return filled.every((t) => isFormTPdfFooterCellText(t));
};

export const extractFormTSEKarnatakaPdfFooter = (rows = [], metaLines = []) => {
  let date = FORM_T_KA_PDF_DATE_LABEL;
  let signatory = FORM_T_KA_PDF_SIGNATORY_LABEL;
  const consider = (text) => {
    const t = normalizeText(text);
    if (!t) return;
    if (isFormTPdfDateFooterText(t)) {
      const rest = t.replace(/^date\s*:?\s*/i, '').trim();
      date = rest ? `Date: ${rest}` : FORM_T_KA_PDF_DATE_LABEL;
    }
    if (isFormTPdfSignatoryFooterText(t) && /authori[sz]ed\s+signatory/i.test(t)) {
      signatory = FORM_T_KA_PDF_SIGNATORY_LABEL;
    }
  };
  (metaLines || []).forEach(consider);
  (rows || []).forEach((row) => {
    (Array.isArray(row) ? row : [row]).forEach(consider);
  });
  return { date, signatory };
};

/** Official Form T: A–I identity, attendance starts at column J (0-based 9). */
export const FORM_T_KA_PDF_ATTENDANCE_START_COL0 = 9;

const cellAt = (row, col) => normalizeText(Array.isArray(row) ? row[col] : '');

const looksLikePdfEmployeeName = (text) => {
  const t = normalizeText(text);
  if (!t || t.length < 2 || t.length > 80) return false;
  if (/^\d+([./-]\d+)*$/.test(t)) return false;
  if (/^(male|female|p|a|wo|h|l|cl|el|sl|od|wop|nil|date:?)$/i.test(t)) return false;
  if (/s\.?\s*no|name of|father|husband|designation|attendance|wages fixed|esi no|uan no|principal employer/i.test(t)) {
    return false;
  }
  return /[a-z]/i.test(t);
};

const isFormTPdfHeaderOrIndexRow = (row) => {
  if (isFormTSEKarnatakaTableHeaderRow(row)) return true;
  if (looksLikePdfEmployeeName(cellAt(row, 1))) return false;
  const filled = (Array.isArray(row) ? row : []).map(normalizeText).filter(Boolean);
  if (filled.filter((c) => /^\d{1,2}$/.test(c)).length >= 8) return true;
  return false;
};

export const formTPdfRowLooksIdentityShiftedToJ = (
  row,
  shift = FORM_T_KA_PDF_ATTENDANCE_START_COL0
) => {
  if (!Array.isArray(row) || isFormTPdfHeaderOrIndexRow(row)) return false;
  const serialA = cellAt(row, 0);
  const nameB = cellAt(row, 1);
  const genderD = cellAt(row, 3);
  const serialJ = cellAt(row, shift);
  const nameK = cellAt(row, shift + 1);
  const genderM = cellAt(row, shift + 3);
  const personAtK = looksLikePdfEmployeeName(nameK);
  const personAtB = looksLikePdfEmployeeName(nameB);
  if (personAtK && !personAtB) return true;
  if (
    /^\d{1,4}$/.test(serialJ) &&
    !/^\d{1,4}$/.test(serialA) &&
    /^(male|female)$/i.test(genderM) &&
    !/^(male|female)$/i.test(genderD)
  ) {
    return true;
  }
  if (
    /^\d{1,4}$/.test(serialJ) &&
    !personAtB &&
    (personAtK || /^(male|female)$/i.test(genderM))
  ) {
    return true;
  }
  return false;
};

const shiftPdfRowFromJToA = (row, shift = FORM_T_KA_PDF_ATTENDANCE_START_COL0) => {
  const src = Array.isArray(row) ? row : [];
  const out = [];
  for (let c = 0; c < src.length; c += 1) {
    out[c] = src[c + shift] != null ? src[c + shift] : '';
  }
  return out;
};

/** Move S.NO/Name from column J onto column A when the Excel dump landed under attendance. */
export const shiftFormTSEKarnatakaPdfIdentityFromJToA = (
  rows,
  shift = FORM_T_KA_PDF_ATTENDANCE_START_COL0
) => {
  const src = Array.isArray(rows) ? rows : [];
  if (!src.length) return src;
  return src.map((row) =>
    formTPdfRowLooksIdentityShiftedToJ(row, shift) ? shiftPdfRowFromJToA(row, shift) : row
  );
};

const findTableStart = (rows) => {
  const src = Array.isArray(rows) ? rows : [];
  for (let r = 0; r < src.length; r += 1) {
    if (isFormTSEKarnatakaTableHeaderRow(src[r])) return r;
  }
  for (let r = 0; r < src.length; r += 1) {
    const filled = (src[r] || []).map((c) => String(c || '').trim()).filter(Boolean);
    if (filled.filter((c) => /^\d{1,2}$/.test(c)).length >= 20) {
      return Math.max(0, r - 1);
    }
  }
  return 0;
};

export const detectFormTSEKarnatakaAttendanceBand = (
  rows,
  tableStart = 0,
  headerBandEnd = 4,
  colCount = 40
) => {
  const src = Array.isArray(rows) ? rows : [];
  const cols = Math.max(1, Number(colCount) || 40);
  const scanFrom = Math.max(0, Number(tableStart) || 0);
  const scanTo = Math.min(Math.max(scanFrom, Number(headerBandEnd) || scanFrom + 4), src.length - 1);
  let dayStart = FORM_T_KA_PDF_ATTENDANCE_START_COL0;
  let dayEnd = -1;
  let dayRow = -1;
  for (let r = scanFrom; r <= scanTo; r += 1) {
    const row = src[r] || [];
    let start = -1;
    const searchFrom = Math.max(0, FORM_T_KA_PDF_ATTENDANCE_START_COL0 - 1);
    for (let c = searchFrom; c < cols; c += 1) {
      if (String(row[c] || '').trim() === '1') {
        start = c;
        break;
      }
    }
    if (start < 0) continue;
    let end = start;
    for (let n = 1; n <= 31; n += 1) {
      const v = String(row[start + n - 1] || '').trim();
      if (v === String(n)) {
        end = start + n - 1;
        continue;
      }
      if (!v && n > 1) continue;
      break;
    }
    if (end - start + 1 >= 20) {
      dayRow = r;
      dayStart = start;
      dayEnd = end;
      break;
    }
  }
  if (dayEnd < dayStart) {
    dayStart = FORM_T_KA_PDF_ATTENDANCE_START_COL0;
    dayEnd = FORM_T_KA_PDF_ATTENDANCE_START_COL0 + 30;
  }
  let labelRow = scanFrom;
  for (let r = scanFrom; r <= scanTo; r += 1) {
    const blob = (src[r] || []).join(' ').toLowerCase();
    if (/attendance/.test(blob)) {
      labelRow = r;
      break;
    }
  }
  return {
    labelRow,
    start: dayStart,
    end: dayEnd,
    dayRow,
    label: FORM_T_KA_PDF_ATTENDANCE_LABEL,
  };
};

const isFormTPdfGroupBannerText = (text) => {
  const t = normalizeText(text).toLowerCase();
  if (!t) return false;
  if (/earned\s+wages(\s+and\s+other\s+allowances)?/.test(t)) return true;
  if (/^deductions?$/.test(t)) return true;
  return false;
};

const collectFormTPdfSourceWageLeaves = (src, dataStart, wageStart, maxCol) => {
  const out = Array(maxCol).fill('');
  const last = Math.min(Math.max(0, Number(dataStart) || 0), src.length);
  for (let r = 0; r < last; r += 1) {
    for (let c = wageStart; c < maxCol; c += 1) {
      const raw = normalizeText(src[r]?.[c]);
      if (!raw || /^\d{1,2}$/.test(raw) || /attendance/i.test(raw)) continue;
      if (isFormTPdfGroupBannerText(raw)) continue;
      out[c] = raw;
    }
  }
  return out;
};

export const applyFormTSEKarnatakaOfficialTableHeaders = (rows, colCount) => {
  const src = (Array.isArray(rows) ? rows : []).map((row) =>
    Array.isArray(row) ? [...row] : []
  );
  if (!src.length) {
    return { rows: src, colCount: Math.max(1, Number(colCount) || 1), attendanceBand: null };
  }
  const band = detectFormTSEKarnatakaAttendanceBand(src, 0, 5, colCount);
  const wageStart = band.end + 1;
  const payableStart = wageStart;
  const earnedStart = payableStart + FORM_T_KA_PDF_PAYABLE_OT_HEADERS.length;
  const earnedEnd = earnedStart + FORM_T_KA_PDF_EARNED_WAGE_LEAVES.length - 1;
  const deductionStart = earnedEnd + 1;
  const deductionEnd = deductionStart + FORM_T_KA_PDF_DEDUCTION_LEAVES.length - 1;
  const trailingStart = deductionEnd + 1;
  const wageColCount = FORM_T_KA_PDF_WAGE_LEAVES.length;
  let dataStart = src.length;
  for (let r = 0; r < src.length; r += 1) {
    if (isFormTPdfHeaderOrIndexRow(src[r])) continue;
    const serial = String(src[r]?.[0] || '').trim();
    const name = String(src[r]?.[1] || '').trim();
    if (/^\d{1,4}$/.test(serial) && looksLikePdfEmployeeName(name)) {
      dataStart = r;
      break;
    }
  }
  if (dataStart >= src.length) {
    const named = src.findIndex((row) => looksLikePdfEmployeeName(String(row?.[1] || '')));
    dataStart = named >= 0 ? named : src.length;
  }
  const maxCol = Math.max(
    Number(colCount) || 0,
    band.end + 1,
    trailingStart + FORM_T_KA_PDF_TRAILING_HEADERS.length,
    40
  );
  const sourceLeaves = collectFormTPdfSourceWageLeaves(src, dataStart, wageStart, maxCol);
  const banner = Array(maxCol).fill('');
  FORM_T_KA_PDF_IDENTITY_HEADERS.forEach((header, i) => {
    banner[i] = header;
  });
  banner[band.start] = FORM_T_KA_PDF_ATTENDANCE_LABEL;
  FORM_T_KA_PDF_PAYABLE_OT_HEADERS.forEach((header, i) => {
    banner[payableStart + i] = sourceLeaves[payableStart + i] || header;
  });
  banner[earnedStart] = FORM_T_KA_PDF_EARNED_WAGES_GROUP;
  banner[deductionStart] = FORM_T_KA_PDF_DEDUCTIONS_GROUP;
  FORM_T_KA_PDF_TRAILING_HEADERS.forEach((header, i) => {
    banner[trailingStart + i] = sourceLeaves[trailingStart + i] || header;
  });
  const days = Array(maxCol).fill('');
  for (let d = 0; d <= band.end - band.start; d += 1) {
    days[band.start + d] = String(d + 1);
  }
  FORM_T_KA_PDF_EARNED_WAGE_LEAVES.forEach((header, i) => {
    const col = earnedStart + i;
    days[col] = sourceLeaves[col] || header;
  });
  FORM_T_KA_PDF_DEDUCTION_LEAVES.forEach((header, i) => {
    const col = deductionStart + i;
    days[col] = sourceLeaves[col] || header;
  });
  const index = Array(maxCol).fill('');
  for (let i = 0; i < FORM_T_KA_PDF_IDENTITY_HEADERS.length; i += 1) {
    index[i] = String(i + 1);
  }
  index[band.start] = '10';
  for (let i = 0; i < wageColCount; i += 1) {
    index[wageStart + i] = String(11 + i);
  }
  return {
    rows: [banner, days, index, ...src.slice(dataStart)],
    colCount: maxCol,
    attendanceBand: {
      labelRow: 0,
      start: band.start,
      end: band.end,
      label: FORM_T_KA_PDF_ATTENDANCE_LABEL,
    },
  };
};

export const normalizeFormTSEKarnatakaPdfMatrix = (
  rows,
  colCount,
  tableStartRow = 0,
  incomingMetaLines = []
) => {
  const src = Array.isArray(rows) ? rows : [];
  const detectedStart = findTableStart(src);
  const start = detectedStart > 0 ? detectedStart : Math.max(0, Number(tableStartRow) || 0);
  const preRows = src.slice(0, start);
  const formTKAFooter = extractFormTSEKarnatakaPdfFooter(src, incomingMetaLines);
  const tableRows = shiftFormTSEKarnatakaPdfIdentityFromJToA(
    src.slice(start).filter((row) => !isFormTSEKarnatakaPdfFooterOnlyRow(row))
  );
  const official = applyFormTSEKarnatakaOfficialTableHeaders(tableRows, colCount);
  const finalTableRows = official.rows;
  const attendanceBand = official.attendanceBand;
  const metaFromRows = preRows
    .map((row) =>
      (Array.isArray(row) ? row : [])
        .map(normalizeText)
        .filter(Boolean)
        .join(' ')
    )
    .filter(Boolean)
    .filter((line) => !isLeakedFormXiiiTitle(line));
  const incoming = (incomingMetaLines || [])
    .map(normalizeText)
    .filter(Boolean)
    .filter((line) => !isLeakedFormXiiiTitle(line));
  const metaLines = [];
  const seen = new Set();
  [...incoming, ...metaFromRows].forEach((line) => {
    const key = line.toLowerCase();
    if (!key || seen.has(key)) return;
    if (isLeakedFormXiiiTitle(line)) return;
    seen.add(key);
    metaLines.push(line);
  });
  const titles = getFormTSEKarnatakaHeaderTitles(metaLines, src, start);
  const fields = extractFormTSEKarnatakaHeaderFields(metaLines, src, start);
  const orderedMeta = [];
  [...titles, ...fields].forEach((line) => {
    const key = normalizeText(line).toLowerCase();
    if (!key || orderedMeta.some((x) => x.toLowerCase() === key)) return;
    orderedMeta.push(line);
  });
  metaLines.forEach((line) => {
    if (isTitleOrCitationLine(line) || isFormTSEKarnatakaAdminRowBlob(line)) return;
    const key = line.toLowerCase();
    if (orderedMeta.some((x) => x.toLowerCase() === key)) return;
    orderedMeta.push(line);
  });

  let maxCol = Math.max(1, Number(official.colCount) || Number(colCount) || 0);
  finalTableRows.forEach((row) => {
    if (!Array.isArray(row)) return;
    maxCol = Math.max(maxCol, row.length);
  });

  return {
    rows: finalTableRows.length ? finalTableRows : src,
    colCount: maxCol,
    tableStartRow: 0,
    metaLines: orderedMeta,
    formTKALayout: true,
    attendanceBand,
    formTKAFooter,
  };
};

export const applyFormTSEKarnatakaPdfNormalization = (target, normalized) => {
  if (!target || !normalized) return target;
  target.rows = normalized.rows;
  target.colCount = normalized.colCount;
  target.tableStartRow = normalized.tableStartRow;
  if (Array.isArray(normalized.metaLines) && normalized.metaLines.length) {
    target.metaLines = normalized.metaLines;
  }
  target.formTKALayout = true;
  if (normalized.attendanceBand) target.attendanceBand = normalized.attendanceBand;
  if (normalized.formTKAFooter) target.formTKAFooter = normalized.formTKAFooter;
  return target;
};
