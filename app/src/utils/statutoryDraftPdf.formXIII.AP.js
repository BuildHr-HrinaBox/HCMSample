import {
  collectUniqueApHeaderLines,
  pickExclusiveApHeaderTitles
} from './statutoryDraftPdf.apHeaderTitles';
import { looksLikeFormTSEKarnatakaPdfContext } from './statutoryDraftPdf.formT.KA';

const normalizeFormXIIIText = (value) => String(value || '').replace(/\s+/g, ' ').trim();

const FORM_XIII_FIELD_SPECS = [
  {
    key: 'contractor',
    label: 'Name and Address of Contractor.',
    pattern: /name\s+and\s+address\s+of\s+contractor\.?/i,
  },
  {
    key: 'establishment',
    label: 'Name and address of Establishment in/ under which contract is carried on',
    pattern: /name\s+and\s+address\s+of\s+establ(?:ishment|ishemnt)\s+in\s*\/\s*under\s+which\s+contract\s+is\s+carried\s+on/i,
  },
  {
    key: 'nature',
    label: 'Nature and location of work.',
    pattern: /nature\s+and\s+location\s+of\s+work\.?/i,
  },
  {
    key: 'employer',
    label: 'Name and address of Principal Employer',
    pattern: /name\s+and\s+address\s+of\s+principal\s+employer\s*:?/i,
  },
];

export const looksLikeFormXIIIAPPdfContext = (
  metaLines,
  rows,
  sheetName = '',
  fileName = ''
) => {
  if (looksLikeFormTSEKarnatakaPdfContext(metaLines, rows, sheetName, fileName)) {
    return false;
  }
  const blob = [
    ...(metaLines || []),
    ...(rows || []).slice(0, 16).flat(),
    sheetName || '',
    fileName || '',
  ]
    .map(normalizeFormXIIIText)
    .join(' ')
    .toLowerCase();
  if (!blob) return false;
  // Numbered Form T "in lieu of" citations mention Form XIII — those are not this form.
  if (/combined\s+muster\s+roll\s+cum\s+register\s+of\s+wages/.test(blob)) return false;
  if (/rule\s+24\s*\(\s*9[\s-]*b/.test(blob)) return false;
  const hasFormXiiiHeading =
    /(?:^|[^0-9.\s])form\s*[-–]?\s*xiii\b/.test(blob) ||
    /^form\s*[-–]?\s*xiii\b/.test(blob.trim());
  const hasNumberedInLieuCitation = /\d+\.\s*form\s*xiii\s+of\s+rules?\s+\d+/.test(blob);
  return (
    (hasFormXiiiHeading && !hasNumberedInLieuCitation) ||
    (/register\s+of\s+workmen\s+employed\s+by\s+contractor/.test(blob) &&
      /vide\s+rule\s*75/.test(blob))
  );
};

export const isFormXIIIPdfTitleRow = (row = []) => {
  const blob = (Array.isArray(row) ? row : [])
    .map(normalizeFormXIIIText)
    .filter(Boolean)
    .join(' ');
  return (
    /form\s*[-–]?\s*xiii\b/.test(blob) ||
    /register\s+of\s+workmen\s+employed\s+by\s+contractor/.test(blob) ||
    /vide\s+rule\s*75/.test(blob)
  );
};

const collectFormXIIILines = (metaLines, rows, tableStart) => [
  ...(rows || []).slice(0, Math.min(Number(tableStart) || 0, rows?.length || 0)).flat(),
  ...(metaLines || []),
]
  .map(normalizeFormXIIIText)
  .filter(Boolean);

const splitFormXIIIField = (text, spec) => {
  const match = normalizeFormXIIIText(text).match(spec.pattern);
  if (!match) return null;
  return normalizeFormXIIIText(text.slice(match.index + match[0].length).replace(/^\s*:\s*/, ''));
};

const extractFormXIIIFields = (lines) => {
  const values = Object.fromEntries(FORM_XIII_FIELD_SPECS.map(({ key }) => [key, '']));
  let activeKey = '';
  lines.forEach((line) => {
    const spec = FORM_XIII_FIELD_SPECS.find((item) => item.pattern.test(line));
    if (spec) {
      activeKey = spec.key;
      const inline = splitFormXIIIField(line, spec);
      if (inline) values[spec.key] = inline;
      return;
    }
    if (activeKey && !values[activeKey] && !FORM_XIII_FIELD_SPECS.some((item) => item.pattern.test(line))) {
      values[activeKey] = line;
      activeKey = '';
    }
  });
  return values;
};

export const getFormXIIIAPHeaderTitles = (metaLines, rows, tableStart) => {
  const limitedRows = (rows || []).slice(0, Math.min(Number(tableStart) || 16, rows?.length || 0));
  return pickExclusiveApHeaderTitles(collectUniqueApHeaderLines(metaLines, limitedRows.length ? limitedRows : rows), {
    form: /form\s*[-–]?\s*xiii\b/i,
    register: /register\s+of\s+workmen/i,
    rule: /vide\s+rule\s*75/i,
    formFallback: 'FORM XIII',
    registerFallback: 'REGISTER OF WORKMEN EMPLOYED BY CONTRACTOR',
    ruleFallback: '[Vide Rule 75 of the Contract Labour (Regulation and Abolition) Central / A.P. Rules]',
    ruleExtract: (text) => String(text || '').match(/\[?Vide\s+Rule\s*75[^\]]*\]?/i)?.[0] || ''
  });
};

export const extractFormXIIIAPAdministrativeRows = (metaLines, rows, tableStart) => {
  const values = extractFormXIIIFields(collectFormXIIILines(metaLines, rows, tableStart));
  if (!values.establishment && values.contractor) {
    values.establishment = values.contractor;
  }
  return [
    [
      `${FORM_XIII_FIELD_SPECS[0].label} :${values.contractor ? ` ${values.contractor}` : ''}`,
      `${FORM_XIII_FIELD_SPECS[1].label} :${values.establishment ? ` ${values.establishment}` : ''}`,
    ],
    [
      `${FORM_XIII_FIELD_SPECS[2].label} :${values.nature ? ` ${values.nature}` : ''}`,
      `${FORM_XIII_FIELD_SPECS[3].label} :${values.employer ? ` ${values.employer}` : ''}`,
    ],
  ];
};