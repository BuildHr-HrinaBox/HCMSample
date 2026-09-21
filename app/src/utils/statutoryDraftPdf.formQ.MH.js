/**
 * Maharashtra Form Q — Muster-Roll Cum Wage Register PDF (Excel model).
 * Matches the downloaded Excel: FORM Q / See Rule 26(1) / MUSTER-ROLL…,
 * Establishment / Employer / Month, then Sr. No → day band → wage/deduction boxes.
 */

function blobText(metaLines, rows, sheetName, fileName) {
  return [
    ...(metaLines || []),
    ...(rows || []).slice(0, 28).flat(),
    sheetName || '',
    fileName || '',
  ]
    .join(' ')
    .toLowerCase();
}

/** Exclude Karnataka appointment order / Gujarat annual return Form Q. */
export function looksLikeFormQMaharashtraPdfContext(
  metaLines,
  rows,
  sheetName = '',
  fileName = ''
) {
  const blob = blobText(metaLines, rows, sheetName, fileName);
  if (/appointment\s+order|rule\s*24\s*\(?\s*9\s*a/i.test(blob)) return false;
  if (/annual\s+return/.test(blob) && /gujarat/.test(blob)) return false;
  if (/\bform\s*[-–.]?\s*p\b/.test(blob) && /gujarat|form[\s._-]*p[\s._-]*gj/.test(blob)) {
    return false;
  }
  const hasFormQ =
    /\bform\s*[-–.]?\s*q\b/.test(blob) ||
    /form[\s._-]*q[\s._-]*(?:maha|mh)|form_q_maharashtra|form_q_-_maharashtra/.test(blob);
  const hasMuster =
    /muster[\s-]*roll/.test(blob) && (/wage\s*register/.test(blob) || /cum\s+wage/.test(blob));
  const hasRule26 = /see\s*rule\s*26/.test(blob);
  const hasFingerprint =
    (/interval\s+for\s+rest/.test(blob) || /date\s+of\s+(the\s+)?month/.test(blob)) &&
    (/full\s+name\s+of\s+the\s+worker/.test(blob) || /name\s+of\s+the\s+worker/.test(blob));
  const fileHint = /form[\s._-]*q/i.test(String(sheetName || '')) || /form[\s._-]*q/i.test(String(fileName || ''));
  if (hasFormQ && (hasMuster || hasRule26 || hasFingerprint)) return true;
  if (fileHint && (hasMuster || hasRule26 || hasFingerprint)) return true;
  if (hasMuster && hasRule26 && hasFingerprint) return true;
  return false;
}

export const FORM_Q_MH_DATE_OF_MONTH_GROUP_LABEL = 'Date of the Month';
export const FORM_Q_MH_PDF_TABLE_FONT_SIZE = 8;

export function isFormQMaharashtraDateOfMonthGroupLabel(text) {
  const t = String(text || '')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
  if (!t) return false;
  if (/date\s+of\s+payment/.test(t) || /entry\s+into\s+service/.test(t)) return false;
  return /^date\s+of\s+(the\s+)?month(?:\s*\(\s*\d+\s*\))?(?:[_\s]+\d{1,2})?$/.test(t);
}

export function isFormQMaharashtraSerialHeaderText(text) {
  const t = String(text || '')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
  return /^(?:sr\.?\s*no|s\.?\s*no|serial\s*(?:no|number))/.test(t);
}

function isDayNumberLeaf(text) {
  const t = String(text || '')
    .replace(/\s+/g, '')
    .trim();
  if (!/^\(?\d{1,2}\)?$/.test(t)) return false;
  const n = Number(String(t).replace(/[()]/g, ''));
  return n >= 1 && n <= 31;
}

function isFormQMaharashtraWageTailHeaderText(text) {
  const t = String(text || '')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
  if (!t) return false;
  return (
    /total\s+days?\s+worked/.test(t) ||
    /minimum\s+rate\s+of\s+wages/.test(t) ||
    /total\s+production/.test(t) ||
    /actual\s+wages?\s+paid/.test(t) ||
    /house\s+rent/.test(t) ||
    /dearness/.test(t) ||
    /gross\s+amount/.test(t) ||
    /overtime/.test(t) ||
    /provident\s+fund/.test(t) ||
    /family\s+pension/.test(t) ||
    /esi\s+contribution/.test(t) ||
    /professional\s+tax/.test(t) ||
    /income\s+tax/.test(t) ||
    /loan.*interest/.test(t) ||
    /^advances?$/.test(t) ||
    /other\s+deductions?/.test(t) ||
    /total\s+deductions?/.test(t) ||
    /net\s+payable/.test(t) ||
    /date\s+of\s+payment/.test(t) ||
    /signature|thumb/.test(t)
  );
}

export function isFormQMaharashtraWorkingHoursGroupLabel(text) {
  const t = String(text || '')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
  return /^working\s+hours/.test(t) && !/overtime/.test(t);
}

export function isFormQMaharashtraIntervalForRestGroupLabel(text) {
  const t = String(text || '')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
  return /interval\s+for\s+rest/.test(t);
}

export function isFormQMaharashtraDeductionsGroupLabel(text) {
  const t = String(text || '')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
  return /^deductions?$/.test(t) || /^deductions?\s*\(/.test(t);
}

function isWorkingHoursGroupLabel(text) {
  return isFormQMaharashtraWorkingHoursGroupLabel(text);
}

function isIntervalForRestGroupLabel(text) {
  return isFormQMaharashtraIntervalForRestGroupLabel(text);
}

function isDeductionsGroupLabel(text) {
  return isFormQMaharashtraDeductionsGroupLabel(text);
}

function isFromToLeaf(text) {
  return /^(from|to)$/i.test(String(text || '').replace(/\s+/g, ' ').trim());
}

function isDeductionLeafHeader(text) {
  const t = String(text || '')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
  if (!t) return false;
  if (/total\s+deductions?/.test(t)) return false;
  return (
    /provident\s+fund|^pf\b/.test(t) ||
    /family\s+pension/.test(t) ||
    /\besi\b/.test(t) ||
    /professional\s+tax|^pt\b/.test(t) ||
    /income\s+tax|^tds\b/.test(t) ||
    /^loans?\b|loan.*interest/.test(t) ||
    /^advances?\b/.test(t) ||
    /other\s+deductions?/.test(t)
  );
}

function isFormQMhTotalDeductionOrNetHeader(text) {
  const t = String(text || '')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
  return /total\s+deductions?/.test(t) || /net\s+payable/.test(t);
}

function isNonDeductionWageTail(text) {
  const t = String(text || '')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
  return (
    /total\s+days?\s+worked/.test(t) ||
    /minimum\s+rate\s+of\s+wages/.test(t) ||
    /total\s+production/.test(t) ||
    /actual\s+wages?\s+paid/.test(t) ||
    /house\s+rent/.test(t) ||
    /dearness/.test(t) ||
    /gross\s+amount/.test(t) ||
    /overtime/.test(t) ||
    /total\s+deductions?/.test(t) ||
    /net\s+payable/.test(t) ||
    /date\s+of\s+payment/.test(t) ||
    /signature|thumb/.test(t)
  );
}

/** Prefer leaf row under the parent merge (Excel row 10); fall back within header band. */
function readHeaderLeaf(rows, labelRow, headerBandEnd, col) {
  const from = Math.min(labelRow + 1, headerBandEnd);
  const to = Math.max(headerBandEnd, labelRow);
  for (let r = from; r <= to && r < rows.length; r += 1) {
    const t = String(rows[r]?.[col] || '')
      .replace(/\s+/g, ' ')
      .trim();
    if (!t) continue;
    if (isFormQMaharashtraDateOfMonthGroupLabel(t)) continue;
    if (isWorkingHoursGroupLabel(t) || isIntervalForRestGroupLabel(t) || isDeductionsGroupLabel(t)) {
      continue;
    }
    return t;
  }
  // Also peek the label row itself for From/To when parent+leaf share a row in sparse matrices.
  const same = String(rows[labelRow]?.[col] || '')
    .replace(/\s+/g, ' ')
    .trim();
  if (same && (isFromToLeaf(same) || isDayNumberLeaf(same) || isDeductionLeafHeader(same))) {
    return same;
  }
  return '';
}

/**
 * Expand a parent group across empty merge slaves + leaf From/To / day / deduction cells.
 * Always returns at least start+fallbackSpan when leaf pattern matches Excel model.
 */
function expandBandFromLabel(rows, labelRow, headerBandEnd, colCount, bandStart, opts = {}) {
  const {
    isGroupLabel = () => false,
    allowFromTo = false,
    allowDayNumbers = false,
    allowDeductionLeaves = false,
    fallbackSpan = 0,
  } = opts;

  let bandEnd = bandStart;
  let sawValidLeaf = false;

  for (let c = bandStart + 1; c < colCount; c += 1) {
    const group = String(rows[labelRow]?.[c] || '')
      .replace(/\s+/g, ' ')
      .trim();
    if (
      group &&
      !isGroupLabel(group) &&
      !(allowDayNumbers && isDayNumberLeaf(group)) &&
      !(allowFromTo && isFromToLeaf(group))
    ) {
      // Split Excel merges may repeat "Deductions" mid-band — treat as same group.
      if (allowDeductionLeaves && isDeductionsGroupLabel(group)) {
        bandEnd = c;
        continue;
      }
      if (
        isWorkingHoursGroupLabel(group) ||
        isIntervalForRestGroupLabel(group) ||
        isFormQMaharashtraDateOfMonthGroupLabel(group) ||
        isDeductionsGroupLabel(group)
      ) {
        break;
      }
      break;
    }

    const leaf = readHeaderLeaf(rows, labelRow, headerBandEnd, c);
    if (!leaf) {
      // Empty under Total Deduction / Net — never absorb into Deductions.
      if (
        allowDeductionLeaves &&
        (isFormQMhTotalDeductionOrNetHeader(group) ||
          /date\s+of\s+payment|bank\s+account|cheque|amount\s+deposited|signature|thumb/.test(
            group.toLowerCase()
          ))
      ) {
        break;
      }
      // Empty parent merge slave — keep expanding while within fallback or until a stop leaf.
      if (fallbackSpan > 0 && c <= bandStart + fallbackSpan) {
        bandEnd = c;
        continue;
      }
      // Empty day-band cells between day numbers are OK.
      if (allowDayNumbers && sawValidLeaf) {
        // Peek ahead a few cols for another day number before stopping.
        let foundDay = false;
        for (let peek = c; peek <= Math.min(c + 2, colCount - 1); peek += 1) {
          if (isDayNumberLeaf(readHeaderLeaf(rows, labelRow, headerBandEnd, peek))) {
            foundDay = true;
            break;
          }
        }
        if (foundDay) {
          bandEnd = c;
          continue;
        }
      }
      if (allowDeductionLeaves && sawValidLeaf) {
        // Do not walk empty cells past Other Deductions.
        const prevLeaf = readHeaderLeaf(rows, labelRow, headerBandEnd, c - 1);
        if (prevLeaf && /other\s+deductions?/.test(prevLeaf.toLowerCase())) {
          break;
        }
        const peekParent = String(rows[labelRow]?.[c + 1] || '')
          .replace(/\s+/g, ' ')
          .trim();
        if (isFormQMhTotalDeductionOrNetHeader(peekParent) || isFormQMhTotalDeductionOrNetHeader(group)) {
          break;
        }
        bandEnd = c;
        continue;
      }
      break;
    }

    if (isFormQMaharashtraSerialHeaderText(leaf)) break;
    if (isNonDeductionWageTail(leaf) && !allowDeductionLeaves) break;
    if (isNonDeductionWageTail(leaf) && allowDeductionLeaves && !isDeductionLeafHeader(leaf)) {
      // Total Deduction / Net sit outside the Deductions parent merge.
      break;
    }

    // Stray Advances/Other under Total Deduction / Net must not extend Deductions.
    if (
      allowDeductionLeaves &&
      (isFormQMhTotalDeductionOrNetHeader(group) ||
        isFormQMhTotalDeductionOrNetHeader(leaf) ||
        /date\s+of\s+payment|bank\s+account|cheque|amount\s+deposited|signature|thumb/.test(
          group.toLowerCase()
        ))
    ) {
      break;
    }
    if (allowDeductionLeaves && sawValidLeaf) {
      // Stop immediately after the first Other Deductions leaf.
      const prevLeaf = readHeaderLeaf(rows, labelRow, headerBandEnd, c - 1);
      if (
        prevLeaf &&
        /other\s+deductions?/.test(prevLeaf.toLowerCase()) &&
        isDeductionLeafHeader(leaf)
      ) {
        break;
      }
    }

    const okLeaf =
      (allowFromTo && isFromToLeaf(leaf)) ||
      (allowDayNumbers && isDayNumberLeaf(leaf)) ||
      (allowDeductionLeaves && isDeductionLeafHeader(leaf)) ||
      isGroupLabel(leaf);

    if (!okLeaf) break;
    sawValidLeaf = true;
    bandEnd = c;
    if (allowDeductionLeaves && isDeductionLeafHeader(leaf) && /other\s+deductions?/.test(leaf.toLowerCase())) {
      break;
    }
  }

  if (bandEnd > bandStart) return bandEnd;

  // Fallbacks matching Excel: Working hours / Interval = 2 cols; Date = consecutive 1..N days.
  if (allowFromTo && fallbackSpan >= 1) {
    const nextLeaf = readHeaderLeaf(rows, labelRow, headerBandEnd, bandStart + 1);
    if (!nextLeaf || isFromToLeaf(nextLeaf) || isFromToLeaf(String(rows[labelRow]?.[bandStart + 1] || ''))) {
      return Math.min(colCount - 1, bandStart + fallbackSpan);
    }
  }
  if (allowDayNumbers) {
    let end = bandStart;
    for (let c = bandStart; c < colCount && c <= bandStart + 31; c += 1) {
      const leaf = readHeaderLeaf(rows, labelRow, headerBandEnd, c);
      if (c === bandStart && (!leaf || isDayNumberLeaf(leaf) || isFormQMaharashtraDateOfMonthGroupLabel(String(rows[labelRow]?.[c] || '')))) {
        end = c;
        continue;
      }
      if (isDayNumberLeaf(leaf)) {
        end = c;
        continue;
      }
      if (!leaf && end > bandStart) {
        // gap inside day band
        end = c;
        continue;
      }
      break;
    }
    if (end > bandStart) return end;
  }
  if (allowDeductionLeaves && fallbackSpan >= 1) {
    let end = bandStart;
    for (let c = bandStart + 1; c < colCount && c <= bandStart + 10; c += 1) {
      const leaf = readHeaderLeaf(rows, labelRow, headerBandEnd, c);
      if (!leaf || isDeductionLeafHeader(leaf)) {
        end = c;
        if (leaf && /other\s+deductions?/.test(leaf.toLowerCase())) break;
        continue;
      }
      break;
    }
    if (end > bandStart) return end;
  }
  return -1;
}

/**
 * Excel-model group merges (row 9 parent / row 10 leaves / row 11 indices):
 * Working hours | Interval for Rest | Date of the Month | Deductions.
 * Searches 2 rows above tableStart so parent labels above Sr. No. still match.
 */
export function detectFormQMaharashtraPdfGroupBands(rows, tableStart, headerBandEnd, colCount) {
  if (!Array.isArray(rows) || colCount < 8) return [];
  const tableStartN = Math.max(0, Number(tableStart) || 0);
  const searchStart = Math.max(0, tableStartN - 2);
  const end = Math.min(rows.length - 1, Math.max(tableStartN + 4, Number(headerBandEnd) || 0));
  const bands = [];

  const findAndPush = (matcher, label, expandOpts) => {
    let labelRow = -1;
    let bandStart = -1;
    for (let r = searchStart; r <= end; r += 1) {
      for (let c = 0; c < colCount; c += 1) {
        if (!matcher(rows[r]?.[c])) continue;
        labelRow = r;
        bandStart = c;
        break;
      }
      if (labelRow >= 0) break;
    }
    if (labelRow < 0) return;
    // Leaf row is usually labelRow+1; include column-index row in headerBandEnd.
    const leafEnd = Math.max(end, labelRow + 2);
    const bandEnd = expandBandFromLabel(rows, labelRow, leafEnd, colCount, bandStart, {
      isGroupLabel: matcher,
      ...expandOpts,
    });
    if (bandEnd < 0) return;
    bands.push({ labelRow, start: bandStart, end: bandEnd, label });
  };

  findAndPush(isWorkingHoursGroupLabel, 'Working hours', {
    allowFromTo: true,
    fallbackSpan: 1,
  });
  findAndPush(isIntervalForRestGroupLabel, 'Interval for Rest', {
    allowFromTo: true,
    fallbackSpan: 1,
  });
  findAndPush(isFormQMaharashtraDateOfMonthGroupLabel, FORM_Q_MH_DATE_OF_MONTH_GROUP_LABEL, {
    allowDayNumbers: true,
    fallbackSpan: 30,
  });
  findAndPush(isDeductionsGroupLabel, 'Deductions', {
    allowDeductionLeaves: true,
    fallbackSpan: 7,
  });

  // Force one Deductions banner spanning PF … Other (absorb split "Deductions" cells).
  // Never include Total Deduction / Net Payable even if stray Advances/Other sit under them.
  const dedIdx = bands.findIndex((b) => isDeductionsGroupLabel(b.label));
  if (dedIdx >= 0) {
    const first = bands[dedIdx];
    const leafEnd = Math.max(end, first.labelRow + 2);
    let endCol = first.start;
    let sawOther = false;
    for (let c = first.start; c < colCount && c <= first.start + 12; c += 1) {
      const parent = String(rows[first.labelRow]?.[c] || '')
        .replace(/\s+/g, ' ')
        .trim();
      const leaf = readHeaderLeaf(rows, first.labelRow, leafEnd, c);
      if (isFormQMhTotalDeductionOrNetHeader(parent) || isFormQMhTotalDeductionOrNetHeader(leaf)) {
        break;
      }
      if (
        /date\s+of\s+payment|bank\s+account|cheque|amount\s+deposited|signature|thumb/.test(
          parent.toLowerCase()
        )
      ) {
        break;
      }
      if (c === first.start) {
        endCol = c;
        continue;
      }
      if (sawOther) break;
      if (isDeductionsGroupLabel(parent) || isDeductionLeafHeader(leaf) || (!leaf && !parent)) {
        if (!leaf && !parent) {
          const peekLeaf = readHeaderLeaf(rows, first.labelRow, leafEnd, c + 1);
          const peekParent = String(rows[first.labelRow]?.[c + 1] || '').trim();
          if (
            isFormQMhTotalDeductionOrNetHeader(peekParent) ||
            isFormQMhTotalDeductionOrNetHeader(peekLeaf)
          ) {
            break;
          }
          if (
            isDeductionLeafHeader(peekLeaf) ||
            isDeductionsGroupLabel(peekParent) ||
            (!peekLeaf && isDeductionsGroupLabel(String(rows[first.labelRow]?.[c + 2] || '')))
          ) {
            endCol = c;
            continue;
          }
          break;
        }
        endCol = c;
        if (isDeductionLeafHeader(leaf) && /other\s+deductions?/.test(leaf.toLowerCase())) {
          sawOther = true;
        }
        continue;
      }
      if (isNonDeductionWageTail(leaf)) break;
      break;
    }
    bands[dedIdx] = { ...first, end: endCol, label: 'Deductions' };
    return bands.filter((b, i) => i === dedIdx || !isDeductionsGroupLabel(b.label));
  }

  return bands;
}

/** True for any Form Q MH parent group banner (hide on leaf/index rows in PDF). */
export function isFormQMaharashtraParentGroupLabel(text) {
  return (
    isFormQMaharashtraWorkingHoursGroupLabel(text) ||
    isFormQMaharashtraIntervalForRestGroupLabel(text) ||
    isFormQMaharashtraDateOfMonthGroupLabel(text) ||
    isFormQMaharashtraDeductionsGroupLabel(text)
  );
}

/**
 * Heading-box labels that read bottom→top (Excel textRotation 90 / PDF angle 90).
 * Identity + wage/deduction leaves — not group banners, From/To, or day numbers.
 */
export function isFormQMaharashtraVerticalHeaderText(text) {
  const t = String(text || '')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
  if (!t) return false;
  if (isFormQMaharashtraParentGroupLabel(t)) return false;
  if (isFromToLeaf(t) || isDayNumberLeaf(t)) return false;
  if (/^\(?\d{1,3}\)?$/.test(t)) return false;
  if (isFormQMaharashtraSerialHeaderText(t)) return true;
  if (/full\s+name\s+of\s+the\s+worker|name\s+of\s+the\s+worker/.test(t)) return true;
  if (/^designation|nature\s+of\s+work/.test(t)) return true;
  if (/^age$/.test(t) || /^sex$/.test(t)) return true;
  if (/entry\s+into\s+service/.test(t)) return true;
  if (isDeductionLeafHeader(t)) return true;
  if (isFormQMaharashtraWageTailHeaderText(t)) return true;
  return false;
}

/**
 * Split long vertical heading labels into at most 2 lines (word-aware).
 * Short labels stay on one line.
 */
export function splitFormQMaharashtraVerticalHeaderLines(text, maxLines = 2) {
  const label = String(text || '')
    .replace(/\s+/g, ' ')
    .trim();
  if (!label) return [];
  const limit = Math.max(1, Math.min(3, Number(maxLines) || 2));
  if (limit === 1 || label.length <= 16 || !/\s/.test(label)) return [label];

  // Prefer known natural breaks for Form Q identity / wage headers.
  const preferred = [
    [/^full\s+name\s+of\s+the\s+worker$/i, ['Full Name of', 'the worker']],
    [/^name\s+of\s+the\s+worker$/i, ['Name of', 'the worker']],
    [
      /^designation\s+of\s+the\s+worker\s+and\s+nature\s+of\s+work$/i,
      ['Designation of the worker', 'and nature of work'],
    ],
    [/^date\s+of\s+entry\s+into\s+service$/i, ['Date of entry', 'into service']],
    [
      /^total\s+production\s+in\s+case\s+of\s+piece\s+rate/i,
      ['Total production in case', 'of piece rate Rs.'],
    ],
    [
      /^minimum\s+rate\s+of\s+wages\s+payable/i,
      ['Minimum rate of', 'wages payable Rs.'],
    ],
    [
      /^total\s+hours\s+of\s+overtime\s+worked/i,
      ['Total hours of overtime', 'worked during the month'],
    ],
    [
      /^signature\s*\/?\s*thumb\s+impression/i,
      ['Signature / Thumb', 'Impression of the worker'],
    ],
    [/^house\s+rent\s+allowance/i, ['House Rent', 'Allowance Paid Rs.']],
    [/^dearness\s+allowance/i, ['Dearness Allowance', 'Paid Rs.']],
    [/^gross\s+amount\s+payable/i, ['Gross Amount', 'Payable Rs.']],
    [/^other\s+deductions?/i, ['Other Deductions', 'Rs. (if any)']],
    [/^provident\s+fund/i, ['Provident Fund', 'Contribution']],
    [/^professional\s+tax/i, ['Professional', 'Tax Rs.']],
    [/^family\s+pension/i, ['Family Pension', 'Rs.']],
    [/^esi\s+contribution/i, ['ESI Contribution', 'Rs.']],
    [/^total\s+days?\s+worked/i, ['Total Days', 'worked']],
    [/^actual\s+wages?\s+paid/i, ['Actual Wages', 'Paid Rs.']],
    [/^total\s+deductions?/i, ['Total Deduction', 'Rs.']],
    [/^net\s+payable/i, ['Net Payable', 'Rs.']],
    [/^date\s+of\s+payment$/i, ['Date of', 'Payment']],
    [/^amount\s+deposited/i, ['Amount', 'Deposited Rs.']],
    [/^loan\s+and\s+interest/i, ['Loan and', 'Interest']],
    [/^overtime\s+earnings/i, ['Overtime', 'earnings Rs.']],
  ];
  for (const [re, parts] of preferred) {
    if (re.test(label) && parts.length <= limit) return parts.slice(0, limit);
  }

  const words = label.split(/\s+/).filter(Boolean);
  if (words.length <= 2) return [label];

  // Balance by character count near the midpoint.
  const total = label.length;
  let best = 1;
  let bestScore = Infinity;
  for (let i = 1; i < words.length; i += 1) {
    const left = words.slice(0, i).join(' ');
    const right = words.slice(i).join(' ');
    if (!right) continue;
    const score = Math.abs(left.length - right.length) + Math.abs(left.length - total / 2) * 0.25;
    if (score < bestScore) {
      bestScore = score;
      best = i;
    }
  }
  return [words.slice(0, best).join(' '), words.slice(best).join(' ')].filter(Boolean);
}

/** Keep name readable; day cells stay narrow — mirrors Excel column proportions. */
export function formQMaharashtraColumnWeight(headerText, maxDataLen = 0) {
  const h = String(headerText || '')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
  const dataLen = Math.max(0, Number(maxDataLen) || 0);
  if (isFormQMaharashtraSerialHeaderText(h)) return 3.4;
  // Vertical name header — column stays narrow; label reads bottom→top (up to 2 lines).
  if (/full\s+name\s+of\s+the\s+worker|name\s+of\s+the\s+worker/.test(h)) return 6.5;
  if (/^designation/.test(h)) return 6.5;
  if (/father|spouse|husband/.test(h)) return 6;
  if (/^age$/.test(h) || /^sex$/.test(h)) return 3.2;
  if (/nature\s+of\s+work/.test(h)) return 6.5;
  if (/entry\s+into\s+service/.test(h)) return 6;
  if (isWorkingHoursGroupLabel(h) || /working\s+hours/.test(h)) return 5.5;
  if (isIntervalForRestGroupLabel(h)) return 5.5;
  if (/^(from|to)$/.test(h)) return 4.5;
  if (isDayNumberLeaf(h) || isFormQMaharashtraDateOfMonthGroupLabel(h)) {
    return Math.max(2, Math.min(2.4, 2 + Math.min(dataLen, 4) * 0.1));
  }
  if (/signature|thumb/.test(h)) return 5.5;
  if (/date\s+of\s+payment/.test(h)) return 5.5;
  if (isFormQMaharashtraWageTailHeaderText(h) || isDeductionsGroupLabel(h)) {
    return Math.max(4.5, Math.min(6.5, 4.5 + Math.min(dataLen, 10) * 0.15));
  }
  return Math.max(4.5, Math.min(8, 4 + Math.min(dataLen, 12) * 0.2));
}

function fieldRank(text) {
  const t = String(text || '')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
  if (/name\s+of\s+(?:the\s+)?establishment/.test(t)) return 0;
  if (/name\s+of\s+(?:the\s+)?employer/.test(t)) return 1;
  if (/^month\b/.test(t)) return 2;
  return 50;
}

function isFormQMhTitleLine(text) {
  const t = String(text || '')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
  if (!t) return false;
  if (/^form\s*[-–.]?\s*q\b/.test(t) && t.length < 40) return true;
  if (/see\s+rules?\s*26/.test(t)) return true;
  if (/muster[\s-]*roll\s+cum\s+wage\s+register/.test(t)) return true;
  return false;
}

function titleRank(text) {
  const t = String(text || '')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
  if (/^form\s*[-–.]?\s*q\b/.test(t)) return 0;
  if (/see\s+rules?\s*26/.test(t)) return 1;
  if (/muster[\s-]*roll/.test(t)) return 2;
  return 9;
}

/**
 * Excel header order: FORM Q / (See Rule 26(1)) / MUSTER-ROLL … then
 * Name of the Establishment, Name of the Employer, Month.
 */
export function rewriteFormQMaharashtraPdfHeader(titles, fields) {
  const pool = [...(Array.isArray(titles) ? titles : []), ...(Array.isArray(fields) ? fields : [])]
    .map((line) => String(line || '').replace(/\s+/g, ' ').trim())
    .filter(Boolean);

  const nextTitles = [];
  const nextFields = [];
  const seen = new Set();
  pool.forEach((line) => {
    const key = line.toLowerCase();
    if (seen.has(key)) return;
    seen.add(key);
    if (isFormQMhTitleLine(line)) {
      nextTitles.push(line);
      return;
    }
    nextFields.push(line);
  });
  nextTitles.sort((a, b) => titleRank(a) - titleRank(b));
  nextFields.sort((a, b) => fieldRank(a) - fieldRank(b));
  return { titles: nextTitles, fields: nextFields };
}

export function trimFormQMaharashtraLeadingBlankPdfColumns(rows, colCount, tableStartRow = 0) {
  if (!Array.isArray(rows) || rows.length === 0 || colCount <= 1) {
    return { rows, colCount };
  }
  const startRow = Math.max(0, Number(tableStartRow) || 0);
  let lead = 0;
  outer: for (let r = startRow; r < Math.min(rows.length, startRow + 8); r += 1) {
    const row = rows[r] || [];
    for (let c = 0; c < colCount; c += 1) {
      if (isFormQMaharashtraSerialHeaderText(row[c])) {
        lead = c;
        break outer;
      }
    }
  }
  if (lead <= 0) return { rows, colCount };
  const nextCount = Math.max(1, colCount - lead);
  const trimmed = rows.map((row) => {
    const src = Array.isArray(row) ? row : [];
    return Array.from({ length: nextCount }, (_, index) => String(src[lead + index] ?? ''));
  });
  return { rows: trimmed, colCount: nextCount };
}

function isFormQMhSignatureHeaderText(text) {
  const t = String(text || '')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
  return /signature|thumb\s+impression/.test(t);
}

function isFormQMhAmountDepositedHeaderText(text) {
  const t = String(text || '')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
  return /amount\s+deposited/.test(t);
}

function readFormQMhHeaderStackText(rows, tableStartRow, col) {
  // Parent group labels often sit 1–2 rows above the Sr. No. / leaf header row.
  const start = Math.max(0, (Number(tableStartRow) || 0) - 2);
  const parts = [];
  for (let r = start; r < Math.min(rows.length, start + 8); r += 1) {
    const t = String(rows[r]?.[col] || '')
      .replace(/\s+/g, ' ')
      .trim();
    if (t) parts.push(t);
  }
  return parts.join(' ');
}

/**
 * Drop trailing empty columns and the duplicate Amount Deposited / Signature pair
 * that appears after an empty column past the first Signature column.
 * Also remove blank columns between Amount Deposited and Signature.
 */
export function dropFormQMaharashtraTrailingDuplicatePdfColumns(
  rows,
  colCount,
  tableStartRow = 0
) {
  if (!Array.isArray(rows) || colCount < 10) return { rows, colCount };
  const start = Math.max(0, Number(tableStartRow) || 0);

  const headerTextAt = (c) => readFormQMhHeaderStackText(rows, start, c);

  // Drop empty columns that sit between Amount Deposited and Signature.
  const dropCols = new Set();
  let amountCol = -1;
  let signatureCol = -1;
  for (let c = 0; c < colCount; c += 1) {
    const stack = headerTextAt(c);
    if (amountCol < 0 && isFormQMhAmountDepositedHeaderText(stack)) amountCol = c;
    if (signatureCol < 0 && isFormQMhSignatureHeaderText(stack)) signatureCol = c;
  }
  if (amountCol >= 0 && signatureCol > amountCol + 1) {
    for (let c = amountCol + 1; c < signatureCol; c += 1) {
      if (!String(headerTextAt(c) || '').trim()) dropCols.add(c);
    }
  }

  let firstSignatureCol = signatureCol;
  if (firstSignatureCol < 0) {
    for (let c = 0; c < colCount; c += 1) {
      if (isFormQMhSignatureHeaderText(headerTextAt(c))) {
        firstSignatureCol = c;
        break;
      }
    }
  }

  let keepUntil = colCount - 1;
  if (firstSignatureCol >= 0) {
    keepUntil = firstSignatureCol;
    for (let c = firstSignatureCol + 1; c < colCount; c += 1) {
      const stack = headerTextAt(c);
      if (!stack) break;
      if (isFormQMhSignatureHeaderText(stack) || isFormQMhAmountDepositedHeaderText(stack)) {
        break;
      }
      keepUntil = c;
    }
  } else {
    while (keepUntil > 0) {
      let any = false;
      for (let r = 0; r < rows.length; r += 1) {
        if (String(rows[r]?.[keepUntil] || '').trim()) {
          any = true;
          break;
        }
      }
      if (any) break;
      keepUntil -= 1;
    }
  }

  for (let c = keepUntil + 1; c < colCount; c += 1) dropCols.add(c);

  const keep = [];
  for (let c = 0; c < colCount; c += 1) {
    if (c > keepUntil) break;
    if (dropCols.has(c)) continue;
    keep.push(c);
  }
  if (keep.length === 0) return { rows, colCount };
  if (keep.length === colCount) {
    // Still coalesce split Deductions labels.
    const nextRows = rows.map((row) => (Array.isArray(row) ? [...row] : []));
    for (let r = start; r < Math.min(nextRows.length, start + 4); r += 1) {
      let seenDeductions = false;
      for (let c = 0; c < colCount; c += 1) {
        if (!isDeductionsGroupLabel(nextRows[r]?.[c])) continue;
        if (!seenDeductions) {
          seenDeductions = true;
          continue;
        }
        nextRows[r][c] = '';
      }
    }
    return { rows: nextRows, colCount };
  }

  const nextRows = rows.map((row) => {
    const src = Array.isArray(row) ? row : [];
    return keep.map((c) => String(src[c] ?? ''));
  });
  const nextCount = keep.length;

  for (let r = start; r < Math.min(nextRows.length, start + 4); r += 1) {
    let seenDeductions = false;
    for (let c = 0; c < nextCount; c += 1) {
      if (!isDeductionsGroupLabel(nextRows[r]?.[c])) continue;
      if (!seenDeductions) {
        seenDeductions = true;
        continue;
      }
      nextRows[r][c] = '';
    }
  }

  return { rows: nextRows, colCount: nextCount };
}

/**
 * Excel vertical merges put Total Deduction / Net on the parent row while a shifted
 * leaf sometimes still holds Advances / Other Deductions — clear those strays.
 * Also clear a second Advances/Other pair after the first Other Deductions leaf.
 * Parent labels often sit 1–2 rows above tableStart (Sr. No. / leaf row).
 */
export function scrubFormQMaharashtraPdfDeductionLeafDuplicates(
  rows,
  colCount,
  tableStartRow = 0
) {
  if (!Array.isArray(rows) || colCount < 8) return rows;
  const tableStart = Math.max(0, Number(tableStartRow) || 0);
  const searchStart = Math.max(0, tableStart - 2);
  const next = rows.map((row) => (Array.isArray(row) ? [...row] : []));

  // Locate parent / leaf header rows (look above tableStart for group banners).
  let parentRow = searchStart;
  let leafRow = Math.min(searchStart + 1, next.length - 1);
  for (let r = searchStart; r < Math.min(next.length, tableStart + 5); r += 1) {
    const blob = (next[r] || []).join(' ').toLowerCase();
    if (/working\s+hours|interval\s+for\s+rest|date\s+of\s+(the\s+)?month|\bdeductions?\b/.test(blob)) {
      parentRow = r;
      leafRow = r + 1;
      break;
    }
  }
  // Prefer leaf as the first row under parent that holds From/To / day / deduction leaves.
  if (leafRow < next.length) {
    for (let r = parentRow; r <= Math.min(parentRow + 2, next.length - 1); r += 1) {
      const sample = (next[r] || []).slice(0, Math.min(colCount, 80));
      const hasLeaf = sample.some(
        (v) => isFromToLeaf(v) || isDayNumberLeaf(v) || isDeductionLeafHeader(v)
      );
      if (hasLeaf) {
        leafRow = r;
        if (r > parentRow) break;
      }
    }
  }
  if (leafRow >= next.length) leafRow = Math.min(parentRow + 1, next.length - 1);
  if (leafRow === parentRow && parentRow + 1 < next.length) leafRow = parentRow + 1;

  const headerScanFrom = Math.min(parentRow, searchStart);
  const headerScanTo = Math.min(next.length - 1, Math.max(leafRow + 2, tableStart + 3));

  let seenOtherDeductions = false;
  for (let c = 0; c < colCount; c += 1) {
    const parent = String(next[parentRow]?.[c] || '')
      .replace(/\s+/g, ' ')
      .trim();
    const leaf = String(next[leafRow]?.[c] || '')
      .replace(/\s+/g, ' ')
      .trim();
    const stack = readFormQMhHeaderStackText(next, tableStart, c);

    if (isDeductionLeafHeader(leaf) && /other\s+deductions?/.test(leaf.toLowerCase())) {
      seenOtherDeductions = true;
    }

    const isWageTailCol =
      isFormQMhTotalDeductionOrNetHeader(parent) ||
      isFormQMhTotalDeductionOrNetHeader(stack) ||
      /date\s+of\s+payment|bank\s+account|cheque|amount\s+deposited|signature|thumb/.test(
        `${parent} ${stack}`.toLowerCase()
      );

    // Parent/stack is Total Deduction / Net / payment tail — clear stray deduction leaves.
    if (isWageTailCol) {
      for (let r = headerScanFrom; r <= headerScanTo; r += 1) {
        const t = String(next[r]?.[c] || '')
          .replace(/\s+/g, ' ')
          .trim();
        if (!t) continue;
        if (isFormQMhTotalDeductionOrNetHeader(t)) continue;
        if (
          /date\s+of\s+payment|bank\s+account|cheque|amount\s+deposited|signature|thumb/.test(
            t.toLowerCase()
          )
        ) {
          continue;
        }
        if (isDeductionLeafHeader(t) || /^advances?\b/i.test(t) || /other\s+deductions?/i.test(t)) {
          next[r][c] = '';
        } else if (
          /total\s+deductions?|net\s+payable/i.test(t) &&
          (/advances?/i.test(t) || /other\s+deductions?/i.test(t))
        ) {
          // Same cell holds both labels — keep only Total Deduction / Net.
          next[r][c] = /net\s+payable/i.test(t) ? 'Net Payable Rs.' : 'Total Deduction Rs.';
        }
      }
    }

    // After the first Other Deductions, wipe duplicate Advances / Other leaves.
    if (
      seenOtherDeductions &&
      leaf &&
      isDeductionLeafHeader(leaf) &&
      (/^advances?\b/i.test(leaf) || /other\s+deductions?/i.test(leaf))
    ) {
      // Keep the first Other; clear subsequent Advances/Other only when this is past Other.
      if (!/other\s+deductions?/i.test(leaf) || seenOtherDeductions) {
        // If this column's leaf is the first Other, don't clear it — seenOther was set on this col.
        const isFirstOther =
          /other\s+deductions?/i.test(leaf) &&
          !isFormQMhTotalDeductionOrNetHeader(parent) &&
          !isFormQMhTotalDeductionOrNetHeader(stack);
        if (!isFirstOther) {
          next[leafRow][c] = '';
        }
      }
    }
  }

  // Final pass: any column whose header stack mixes Total Deduction/Net with Advances/Other.
  for (let c = 0; c < colCount; c += 1) {
    const stack = readFormQMhHeaderStackText(next, tableStart, c).toLowerCase();
    if (!/total\s+deductions?|net\s+payable/.test(stack)) continue;
    if (!/advances?|other\s+deductions?/.test(stack)) continue;
    for (let r = headerScanFrom; r <= headerScanTo; r += 1) {
      const t = String(next[r]?.[c] || '')
        .replace(/\s+/g, ' ')
        .trim();
      if (!t) continue;
      if (isFormQMhTotalDeductionOrNetHeader(t)) continue;
      if (isDeductionLeafHeader(t) || /^advances?\b/i.test(t) || /other\s+deductions?/i.test(t)) {
        next[r][c] = '';
      }
    }
  }

  return next;
}

export function normalizeFormQMaharashtraPdfMatrix(rows, colCount, tableStartRow = 0, metaLines = []) {
  const trimmed = trimFormQMaharashtraLeadingBlankPdfColumns(rows, colCount, tableStartRow);
  const scrubbedRows = scrubFormQMaharashtraPdfDeductionLeafDuplicates(
    trimmed.rows,
    trimmed.colCount,
    tableStartRow
  );
  const cleaned = dropFormQMaharashtraTrailingDuplicatePdfColumns(
    scrubbedRows,
    trimmed.colCount,
    tableStartRow
  );
  // Scrub again after column drops (indices change).
  const finalRows = scrubFormQMaharashtraPdfDeductionLeafDuplicates(
    cleaned.rows,
    cleaned.colCount,
    tableStartRow
  );
  return {
    rows: finalRows,
    colCount: cleaned.colCount,
    tableStartRow,
    metaLines,
    formQMaharashtraLayout: true,
  };
}
