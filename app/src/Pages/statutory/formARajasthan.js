import {
  formAGJGujaratHeaderNorm,
  formAGJHeaderAliasBucket,
} from './formAGJGujarat';

/** Rajasthan Form A — employee / workman register (Shops & Establishments). */

const FORM_A_RJ_HIDDEN_HEADER_BUCKETS = new Set([
  'employeeId',
  'gender',
  'categoryAddress',
  'bankAccount',
  'bankName',
  'bankBranchIfsc',
]);

export function isFormARajasthanHiddenTableHeader(header) {
  const bucket = formAGJHeaderAliasBucket(formAGJGujaratHeaderNorm(header));
  return FORM_A_RJ_HIDDEN_HEADER_BUCKETS.has(bucket);
}

export function isFormARajasthanContext(
  formHeader,
  rowItem,
  fileName,
  sheetText = '',
  tableHeaders = []
) {
  const parts = [
    rowItem?.formName,
    rowItem?.FormName,
    rowItem?.formFileName,
    rowItem?.FormFileName,
    rowItem?.siteState,
    rowItem?.SiteState,
    rowItem?.state,
    rowItem?.State,
    fileName,
    formHeader?.title,
    formHeader?.subtitle,
    formHeader?.reference,
    sheetText,
    Array.isArray(tableHeaders) ? tableHeaders.join(' ') : '',
  ]
    .filter((x) => x != null && String(x).trim() !== '')
    .join(' ')
    .toLowerCase();

  if (/gujarat|\b_gj\b|form[\s._-]*a[\s._-]*gj|form_a_gj/.test(parts)) return false;
  if (/maternity\s+benefit|register\s+of\s+muster\s+roll/.test(parts)) return false;

  const hasRajasthan =
    /rajasthan|\b_rj\b|form[\s._-]*a[\s._-]*rj|form_a_rj/.test(parts);
  const hasFormA = /\bform[\s._-]*a\b/.test(parts);
  const hasEmployeeWorkmanFormat =
    /format\s+of\s+employee/.test(parts) ||
    (/employee\s*\/\s*workman\s*\/\s*worker/.test(parts) &&
      !/register\s+of\s+wages|rate\s+of\s+wage/.test(parts));

  if (hasRajasthan && (hasFormA || hasEmployeeWorkmanFormat)) return true;
  if (hasRajasthan && /\bform_a_rj\b/.test(parts)) return true;

  return false;
}

export function resolveFormARajasthanTableHeaders(tableHeaders) {
  const parsed = Array.isArray(tableHeaders)
    ? tableHeaders.map((h) => String(h || '').trim()).filter(Boolean)
    : [];
  return parsed.filter((h) => !isFormARajasthanHiddenTableHeader(h));
}

export function remapFormARajasthanRowsToHeaders(rows, sourceHeaders, targetHeaders) {
  const src = Array.isArray(sourceHeaders) ? sourceHeaders : [];
  const tgt = resolveFormARajasthanTableHeaders(targetHeaders);
  if (!Array.isArray(rows)) return [];
  return rows.map((row, rowIndex) => {
    if (!row || typeof row !== 'object') return {};
    const out = {};
    tgt.forEach((targetHeader, colIdx) => {
      let val = '';
      if (Object.prototype.hasOwnProperty.call(row, targetHeader)) {
        val = row[targetHeader];
      } else if (src[colIdx] && Object.prototype.hasOwnProperty.call(row, src[colIdx])) {
        val = row[src[colIdx]];
      } else {
        const bucket = formAGJHeaderAliasBucket(formAGJGujaratHeaderNorm(targetHeader));
        for (const [k, v] of Object.entries(row)) {
          if (formAGJHeaderAliasBucket(formAGJGujaratHeaderNorm(k)) === bucket) {
            val = v;
            break;
          }
        }
      }
      if (
        (val == null || val === '') &&
        formAGJHeaderAliasBucket(formAGJGujaratHeaderNorm(targetHeader)) === 'sno'
      ) {
        val = rowIndex + 1;
      }
      out[targetHeader] = val == null ? '' : val;
    });
    return out;
  });
}
