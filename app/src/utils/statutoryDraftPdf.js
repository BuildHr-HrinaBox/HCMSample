import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import { jsPDF } from 'jspdf';
import * as XLSX from 'xlsx';
import { convertExcelFilesToSmartbrowzPdf } from './smartbrowzPdf';
import {
  FORM_14_RJ_DAY_ENTRIES_NOTE,
  FORM_14_RJ_FOOTER_NOTE
} from '../Pages/statutory/form14Rajasthan';
import { isFormAGJGujaratExportTitleBandText } from '../Pages/statutory/formAGJGujarat';
import {
  FORM_D_RJ_ABSENCE_CODES_NOTE,
  FORM_D_RJ_E_FORM_NOTE,
  FORM_D_RJ_OUTER_FOOTNOTES,
  FORM_D_RJ_RELAY_MINES_NOTE,
  extractFormDRJOuterFootnoteText,
  isFormDRJAbsenceCodesFootnoteText,
  isFormDRJEFormMaintenanceFootnoteText,
  isFormDRJOuterFootnoteRow,
  isFormDRJOuterFootnoteText,
  isFormDRJRelayMinesFootnoteText,
  normalizeFormDRJOuterFootnoteText,
} from '../Pages/statutory/formDRajasthan';
import {
  isFormXXDateOfRecoveryGroupLabel,
  isFormXXAPAdministrativeRow,
  getFormXXAPHeaderTitles,
  isFormXXAPTableHeaderRow,
  looksLikeFormXXAPPdfContext,
} from './statutoryDraftPdf.formXX.AP';
import {
  getFormXXIAPHeaderTitles,
  looksLikeFormXXIAPPdfContext
} from './statutoryDraftPdf.formXXI.AP';
import {
  formXVIIAPHeaderTextMetrics,
  getFormXVIIAPHeaderTitles,
  isFormXVIIAdministrativeRow,
  isFormXVIITableHeaderRow,
  looksLikeFormXVIIAPPdfContext
} from './statutoryDraftPdf.formXVII.AP';
import {
  looksLikeFormXIXMPCLRAPdfContext,
  normalizeFormXIXMPCLRAWageSlipPdfMatrix
} from './stautoryDraftPdf.formXIX.MP.CLRA';
import {
  looksLikeFormXIXGJWageSlipPdfContext,
  normalizeFormXIXGJWageSlipPdfMatrix
} from './statutoryDraftPdf.formXIX.GJ';
import {
  looksLikeFormXIXKarnatakaPdfContext,
  normalizeFormXIXKarnatakaWageSlipPdfMatrix
} from './statutoryDraftPdf.formXIX.KA';
import {
  looksLikeFormXIXTamilNaduPdfContext,
  resolveFormXIXTamilNaduNetAmountPdfSpan,
} from './statutoryDraftPdf.formXIX.TN';
import {
  looksLikeFormXIVKarnatakaPdfContext,
  normalizeFormXIVKarnatakaPdfMatrix
} from './statutoryDraftPdf.formXIV.KA';
import {
  looksLikeFormKGJGujaratPdfContext,
  rewriteFormKGJGujaratPdfHeader,
  sanitizeFormKGJGujaratPdfDataRows,
  trimFormKGJGujaratLeadingBlankPdfColumns,
} from './statutoryDraftPdf.formKGJ.GJ';
import {
  detectFormLGJGujaratPdfGroupBands,
  formLGJGujaratColumnWeight,
  isFormLGJDateOfMonthGroupLabel,
  isFormLGJWeeklyHolidayHeaderText,
  looksLikeFormLGJGujaratPdfContext,
  sanitizeFormLGJGujaratPdfHeaderRows,
  trimFormLGJGujaratLeadingBlankPdfColumns,
} from './statutoryDraftPdf.formLGJ.GJ';
import {
  detectFormPGJGujaratPdfGroupBands,
  formPGJGujaratColumnWeight,
  isFormPGJDateOfMonthGroupLabel,
  looksLikeFormPGJGujaratPdfContext,
  normalizeFormPGJGujaratPdfMatrix,
  dropFormPGJGujaratExtraDayPdfColumns,
  blankFormPGJGujaratMinimumRateOfWagesPdfCells,
  resolveFormPGJGujaratPdfDaysInMonth,
  rewriteFormPGJGujaratPdfHeader,
  trimFormPGJGujaratLeadingBlankPdfColumns,
} from './statutoryDraftPdf.formPGJ.GJ';
import {
  FORM_Q_MH_PDF_TABLE_FONT_SIZE,
  detectFormQMaharashtraPdfGroupBands,
  formQMaharashtraColumnWeight,
  isFormQMaharashtraDateOfMonthGroupLabel,
  isFormQMaharashtraParentGroupLabel,
  isFormQMaharashtraVerticalHeaderText,
  looksLikeFormQMaharashtraPdfContext,
  normalizeFormQMaharashtraPdfMatrix,
  rewriteFormQMaharashtraPdfHeader,
  splitFormQMaharashtraVerticalHeaderLines,
} from './statutoryDraftPdf.formQ.MH';
import {
  looksLikeFormOGJGujaratPdfContext,
  rewriteFormOGJGujaratPdfHeader,
  trimFormOGJGujaratLeadingBlankPdfColumns,
} from './statutoryDraftPdf.formOGJ.GJ';
import {
  FORM_XVIII_TN_PDF_TABLE_FONT_SIZE,
  coalesceFormXVIIITamilNaduTitleLines,
  detectFormXVIIITamilNaduPdfGroupBands,
  isFormXVIIITamilNaduAmountOfWagesEarnedGroupLabel,
  isFormXVIIITamilNaduWagesCumMusterTitle,
  looksLikeFormXVIIITamilNaduPdfContext,
} from '../Pages/statutory/formXVIIITamilNaduWagesMuster';
import {
  detectFormBTamilNaduPdfGroupBands,
  formBTamilNaduPdfColumnWeight,
  isFormBTamilNaduAmountsDeductedGroupLabel,
  isFormBTamilNaduPdfSignatoryRow,
  isFormBTamilNaduPdfSignatoryText,
  looksLikeFormBTamilNaduPdfContext,
  normalizeFormBTamilNaduPdfMatrix,
  rewriteFormBTamilNaduPdfHeader,
} from '../Pages/statutory/formBTamilNadu';
import {
  isFormCTamilNaduLwfPdfSignatoryRow,
  isFormCTamilNaduLwfPdfSignatoryText,
  rewriteFormCTamilNaduLwfPdfHeader,
} from '../Pages/statutory/formCTamilNaduLwf';
import {
  getFormXVIAPHeaderTitles,
  isFormXVIIDateNumberRow,
  isFormXVIAPTableHeaderRow,
  looksLikeFormXVIAPPdfContext,
  sanitizeFormXVIAPMonthValue,
  sanitizeFormXVIAPNatureValue
} from './statutoryDraftPdf.formXVI.AP';
import {
  extractFormXIIIAPAdministrativeRows,
  getFormXIIIAPHeaderTitles,
  isFormXIIIPdfTitleRow,
  looksLikeFormXIIIAPPdfContext
} from './statutoryDraftPdf.formXIII.AP';
import {
  extractFormXXIIIAPSEAdminLayout,
  findFormXXIIIAPSEDeductionsBand,
  FORM_XXIII_APSE_ADMIN_WIDTHS,
  formXXIIIAPSEColumnWeight,
  formatFormXXIIIAPSEMoneyPdfValue,
  getFormXXIIIAPSEHeaderTitles,
  isFormXXIIIAPSEAdminRow,
  isFormXXIIIAPSECenterValueHeader,
  isFormXXIIIAPSEDeductionsGroupLabel,
  isFormXXIIIAPSEMoneyHeader,
  isFormXXIIIAPSETableHeaderRow,
  isFormXXIIIAPSETitleRow,
  looksLikeFormXXIIIAPSEPdfContext,
  trimFormXXIIIAPSEEmptyPdfColumns
} from './statutoryDraftPdf.formXXIII.AP.S&D.js';
import {
  formXXIIIGJOtPdfColumnWeight,
  getFormXXIIIGJOtHeaderTitles,
  isFormXXIIIGJOtAdminRow,
  isFormXXIIIGJOtCenterValueHeader,
  isFormXXIIIGJOtPreambleRow,
  isFormXXIIIGJOtSpuriousPdfDataRow,
  isFormXXIIIGJOtTableHeaderRow,
  isFormXXIIIGJOtTitleRow,
  looksLikeFormXXIIIGJOtPdfContext,
  normalizeFormXXIIIGJOtPdfMatrix,
  sanitizeFormXXIIIGJOtMatrixCell,
} from './statutoryDraftPdf.formXXIII.GJ.js';
import {
  applyFormTSEKarnatakaPdfNormalization,
  detectFormTSEKarnatakaAttendanceBand,
  extractFormTSEKarnatakaHeaderFields,
  FORM_T_KA_PDF_DATE_LABEL,
  FORM_T_KA_PDF_SIGNATORY_LABEL,
  getFormTSEKarnatakaHeaderTitles,
  isFormTSEKarnatakaAdminRowBlob,
  isFormTSEKarnatakaPdfFooterOnlyRow,
  isFormTSEKarnatakaTableHeaderRow,
  looksLikeFormTSEKarnatakaPdfContext,
  normalizeFormTSEKarnatakaPdfMatrix,
} from './statutoryDraftPdf.formT.KA';
import {
  applyFormQKarnatakaPdfNormalization,
  looksLikeFormQKarnatakaPdfContext,
  normalizeFormQKarnatakaPdfMatrix,
} from './statutoryDraftPdf.formQ.KA';
import {
  applyFormMGJGujaratPdfNormalization,
  drawFormMGJGujaratIdentityCard,
  looksLikeFormMGJGujaratPdfContext,
  normalizeFormMGJGujaratPdfMatrix,
} from './statutoryDraftPdf.formMGJ.GJ';
import {
  applyFormNGJGujaratPdfNormalization,
  drawFormNGJGujaratLeaveBook,
  looksLikeFormNGJGujaratPdfContext,
  normalizeFormNGJGujaratPdfMatrix,
} from './statutoryDraftPdf.formNGJ.GJ';
import {
  applyFormFKarnatakaPdfNormalization,
  drawFormFKarnatakaLeaveRegister,
  looksLikeFormFKarnatakaPdfContext,
  normalizeFormFKarnatakaPdfMatrix,
} from './statutoryDraftPdf.formF.KA';
import {
  applyFormPKarnatakaPdfNormalization,
  drawFormPKarnatakaNotice,
  looksLikeFormPKarnatakaPdfContext,
  normalizeFormPKarnatakaPdfMatrix,
} from './statutoryDraftPdf.formP.KA';
import { isApPdfHeadingRow } from './statutoryDraftPdf.apHeaderTitles';
import { buildFormXVAPServiceCertificatePdfBlob } from './statutoryDraftPdf.formXV.AP';
const EXCEL_EXT_RE = /\.(xlsx|xls|xlsm|xlsb)$/i;
/** Form T KA needs ~70 leaf cols (9 identity + 31 days + wage/deduction band). */
const MAX_PDF_COLS = 80;
const MAX_PDF_DATA_ROWS = 2500;
const MAX_ZIP_EXCEL_FILES = 60;
const MAX_TRAILING_EMPTY_AFTER_CONTENT = 2;
const SYSTEM_GENERATED_DOCUMENT_NOTE = 'This is a System Generated Document';

/**
 * True when text looks like a draft Excel/ZIP file base name
 * (e.g. "Form_W_-_TamilNadu"), not a real form heading.
 */
const looksLikeExcelDraftFileLabel = (text) => {
  const raw = String(text || '')
    .replace(EXCEL_EXT_RE, '')
    .replace(/\.zip$/i, '')
    .trim();
  if (!raw) return false;
  // Filename-style separators: Form_W_-_TamilNadu, Form-L-Gujarat.xlsx stem
  if (/_-_/.test(raw) || /__+/.test(raw)) return true;
  if (/^form[_\s.-]+[a-z0-9xivlc.]+[_\s.-]+-+[_\s.-]*[a-z]/i.test(raw)) return true;
  // Underscore-heavy catalog names that are not printed form titles
  const underscores = (raw.match(/_/g) || []).length;
  if (underscores >= 2 && /form/i.test(raw)) return true;
  return false;
};

/**
 * Excel sheet tab labels must not become PDF titles
 * (e.g. "LWF Act - Form C", bare "Form C").
 */
const looksLikeExcelSheetTabName = (text) => {
  const raw = String(text || '')
    .replace(EXCEL_EXT_RE, '')
    .replace(/\.zip$/i, '')
    .trim();
  if (!raw) return false;
  if (looksLikeExcelDraftFileLabel(raw)) return true;
  if (/\bact\s*[-–—]\s*form\b/i.test(raw)) return true;
  if (/^lwf\b/i.test(raw) && /\bform\b/i.test(raw)) return true;
  // Bare short tab names like "Form C" / "Form-C" when used as sheet.name
  if (/^form\s*[-–.]?\s*[a-z0-9xivlc.]+\s*$/i.test(raw) && raw.length <= 24) return true;
  return false;
};

/**
 * Normalize Excel soft-breaks / `_x000d_` and expand a title/meta cell into
 * separate lines (Excel wrapText headings must not become one PDF line).
 */
const normalizeStatutoryMultilineText = (raw) =>
  String(raw || '')
    .replace(/_x000d_/gi, '\n')
    .replace(/\u000d/g, '\n')
    .replace(/\r\n/g, '\n')
    .replace(/\r/g, '\n')
    .replace(/\n{3,}/g, '\n\n')
    .trim();

const TITLE_CONCAT_SPLIT_RE =
  /\s+(?=(?:REGISTER\s+OF|OVERTIME\s+MUSTER\s+ROLL|(?<!OVERTIME\s)MUSTER\s+ROLL|LIST\s+OF|WAGE\s+SLIP|LETTER\s+OF|NOTICE\s+OF|COMBINED\s+|\(Prescribed\s+under\b|\[Prescribed\s+under\b|\[?\(?See\b|THE\s+(?:TAMIL|ANDHRA|KARNATAKA|RAJASTHAN|MADHYA|GUJARAT|TELANGANA|KERALA|WEST\s+BENGAL|ODISHA|PUNJAB|HARYANA|CENTRAL)\b))/i;

const expandStatutoryMetaSegments = (raw) => {
  const normalized = normalizeStatutoryMultilineText(raw);
  if (!normalized) return [];
  let parts = normalized
    .split(/\n+/)
    .map((s) => s.replace(/[ \t\u00a0]+/g, ' ').trim())
    .filter(Boolean);

  // Form XVIII TN subtitle must stay one line (do not split on "Register of").
  const joinedNorm = parts.join(' ').replace(/\s+/g, ' ').trim();
  if (isFormXVIIITamilNaduWagesCumMusterTitle(joinedNorm)) {
    return [joinedNorm];
  }

  // SheetJS / some exports flatten wrapText titles into one long line — re-split.
  if (parts.length === 1) {
    const blob = parts[0];
    const looksLikeStackedTitle =
      /^form\b/i.test(blob) &&
      blob.length > 36 &&
      !isFormXVIIITamilNaduWagesCumMusterTitle(blob) &&
      (TITLE_CONCAT_SPLIT_RE.test(blob) ||
        /register of|see\s+(?:sub-)?rule|prescribed\s+under|overtime\s+muster/i.test(blob));
    if (looksLikeStackedTitle) {
      const split = blob
        .split(TITLE_CONCAT_SPLIT_RE)
        .map((s) => s.replace(/\s+/g, ' ').trim())
        .filter(Boolean);
      if (split.length > 1) parts = split;
    }
  }
  return coalesceFormXVIIITamilNaduTitleLines(parts);
};

const isSystemGeneratedDocumentNote = (text) =>
  /^this\s+is\s+a\s+system\s+generated\s+document\.?$/i.test(
    String(text || '')
      .replace(/\s+/g, ' ')
      .trim()
  );

const isSystemGeneratedDocumentNoteRow = (row) => {
  if (!Array.isArray(row)) return false;
  const filled = row.map((c) => String(c || '').trim()).filter(Boolean);
  return filled.length === 1 && isSystemGeneratedDocumentNote(filled[0]);
};

/** Form C LWF legal footnote under the register grid. */
const FORM_C_UNPAID_ACCUMULATIONS_FOOTNOTE =
  '*See definition of "Unpaid Accumulations" under Section 2(I) of the Tamil Nadu Labour Welfare Fund Act, 1972';
const FORM_C_RJ_FOOTNOTE = '*Applicable only in case of damage/loss/fine';
const FORM_C_GJ_WHEREVER_APPLICABLE_FOOTNOTE = '*Wherever applicable';

/** Gujarat Form D — legal notes under the muster-roll box (outside the border). */
const FORM_DGJ_GJ_ELECTRONIC_FORMAT_NOTE_1 =
  '*Not necessary in case of electronic format';
const FORM_DGJ_GJ_ELECTRONIC_FORMAT_NOTE_2 =
  '**Not necessary in case of electronic format';
const FORM_DGJ_GJ_GOVERNOR_ORDER_NOTE =
  'By order and in the name of the Governor of Gujarat';
const FORM_DGJ_GJ_OUTER_FOOTNOTES = [
  FORM_DGJ_GJ_ELECTRONIC_FORMAT_NOTE_1,
  FORM_DGJ_GJ_ELECTRONIC_FORMAT_NOTE_2,
  FORM_DGJ_GJ_GOVERNOR_ORDER_NOTE,
];

/** Gujarat Form A — legal notes under the employee table (outside the border). */
const FORM_AGJ_GJ_AGE_NOTE =
  '# NOTE : in case the age is between 14 to 18 year, mention the nature of work, daily hours of work and intervals of rest in remarks Column';
const FORM_AGJ_GJ_SKILL_NOTE = '* (Highly Skilled/Semi Skilled/ Un Skilled).';
const FORM_AGJ_GJ_ELECTRONIC_NOTE = '**Not necessary in case of electronic format';
const FORM_AGJ_GJ_WHEREVER_NOTE = '***Wherever applicable';
const FORM_AGJ_GJ_OUTER_FOOTNOTES = [
  FORM_AGJ_GJ_AGE_NOTE,
  FORM_AGJ_GJ_SKILL_NOTE,
  FORM_AGJ_GJ_ELECTRONIC_NOTE,
  FORM_AGJ_GJ_WHEREVER_NOTE,
];

const normalizeFormAGJFootnoteCompare = (text) =>
  String(text || '')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase()
    .replace(/elecrtonic/g, 'electronic');

const isFormAGJGujaratOuterFootnoteText = (text) => {
  const norm = normalizeFormAGJFootnoteCompare(text);
  if (!norm) return false;
  if (/#?\s*note/.test(norm) && /14\s*to\s*18/.test(norm) && /remarks/.test(norm)) return true;
  if (/highly\s*skilled/.test(norm) && /semi\s*skilled/.test(norm) && /un\s*skilled/.test(norm)) {
    return true;
  }
  if (/^\*{1,2}\s*not\s+necessary\s+in\s+case\s+of\s+electronic\s+format\.?$/.test(norm)) {
    return true;
  }
  if (/^\*{1,3}\s*wherever\s+applicable\.?$/.test(norm)) return true;
  return false;
};

const isFormAGJGujaratOuterFootnoteRow = (row) => {
  if (!Array.isArray(row)) return false;
  const filled = row.map((c) => String(c || '').trim()).filter(Boolean);
  if (!filled.length) return false;
  const unique = [...new Set(filled)];
  return unique.length === 1 && isFormAGJGujaratOuterFootnoteText(unique[0]);
};

const extractFormAGJGujaratOuterFootnoteText = (row) => {
  if (!Array.isArray(row)) return '';
  for (let i = 0; i < row.length; i += 1) {
    const t = String(row[i] || '')
      .replace(/\s+/g, ' ')
      .trim();
    if (!isFormAGJGujaratOuterFootnoteText(t)) continue;
    const norm = normalizeFormAGJFootnoteCompare(t);
    if (/14\s*to\s*18/.test(norm)) return FORM_AGJ_GJ_AGE_NOTE;
    if (/highly\s*skilled/.test(norm)) return FORM_AGJ_GJ_SKILL_NOTE;
    if (/not\s+necessary/.test(norm)) return FORM_AGJ_GJ_ELECTRONIC_NOTE;
    if (/wherever\s+applicable/.test(norm)) return FORM_AGJ_GJ_WHEREVER_NOTE;
    return t;
  }
  return '';
};

const looksLikeFormAGJGujaratPdfContext = (metaLines, rows, sheetName = '') => {
  const blob = [...(metaLines || []), ...(rows || []).slice(0, 22).flat(), sheetName || '']
    .join(' ')
    .toLowerCase();
  if (/rajasthan|\b_rj\b|form[\s._-]*a[\s._-]*rj|form_a_rj/.test(blob)) return false;
  // Form B wage register must never inherit Form A footnotes / header pairing.
  if (
    /\bform\s*b\b|form[\s._-]*b[\s._-]|form_b_gj|format\s+of\s+wage\s+register/.test(blob) ||
    (/wage\s+register/.test(blob) && /rate\s+of\s+(?:wage|minimum)/.test(blob))
  ) {
    return false;
  }
  const hasFormA = /\bform\s*a\b|form[\s._-]*a[\s._-]*gj|form_a_gj/.test(blob);
  const hasEmployeeFormat =
    /format\s+of\s+employee/.test(blob) ||
    (/employee\s*\/\s*workman\s*\/\s*worker/.test(blob) &&
      /date\s+of\s+birth|education\s+level|father/.test(blob) &&
      !/register\s+of\s+wages|wage\s+register|rate\s+of\s+wage/.test(blob));
  const hasGujarat = /gujarat|\b_gj\b|form[\s._-]*a[\s._-]*gj|form_a_gj/.test(blob);
  if (hasGujarat && (hasFormA || hasEmployeeFormat)) return true;
  if (hasFormA && hasEmployeeFormat) return true;
  return false;
};

/** Stack Form A GJ header fields: Establishment, then NAME OF OWNER below (not same row). */
const reorderFormAGJPdfHeaderFields = (fields) => {
  const list = Array.isArray(fields) ? fields.map((f) => String(f || '').trim()).filter(Boolean) : [];
  let est = '';
  let owner = '';
  const others = [];
  list.forEach((f) => {
    const t = f.toLowerCase();
    if (!est && /name\s+of\s+establishment/.test(t) && !/principal/.test(t)) {
      est = f;
      return;
    }
    if (!owner && /name\s+of\s+owner/.test(t)) {
      owner = f;
      return;
    }
    others.push(f);
  });
  const out = [];
  if (est) out.push(est);
  if (owner) out.push(owner);
  else if (est) out.push('NAME OF OWNER');
  out.push(...others);
  return out;
};

/** @deprecated Use reorderFormAGJPdfHeaderFields — owner must be below, not beside. */
const buildFormAGJPdfSplitFieldRows = (fields) => {
  const ordered = reorderFormAGJPdfHeaderFields(fields);
  return { splitRows: [], remaining: ordered };
};

const normalizeFormDGJFootnoteCompare = (text) =>
  String(text || '')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase()
    .replace(/elecrtonic/g, 'electronic');

const isFormDGJGujaratOuterFootnoteText = (text) => {
  const norm = normalizeFormDGJFootnoteCompare(text);
  if (!norm) return false;
  if (/^\*{1,2}\s*not\s+necessary\s+in\s+case\s+of\s+electronic\s+format\.?$/.test(norm)) {
    return true;
  }
  if (/^by\s+order\s+and\s+in\s+the\s+name\s+of\s+the\s+governor\s+of\s+gujarat\.?$/.test(norm)) {
    return true;
  }
  return false;
};

const isFormDGJGujaratOuterFootnoteRow = (row) => {
  if (!Array.isArray(row)) return false;
  const filled = row.map((c) => String(c || '').trim()).filter(Boolean);
  if (!filled.length) return false;
  const unique = [...new Set(filled)];
  return unique.length === 1 && isFormDGJGujaratOuterFootnoteText(unique[0]);
};

const extractFormDGJGujaratOuterFootnoteText = (row) => {
  if (!Array.isArray(row)) return '';
  for (let i = 0; i < row.length; i += 1) {
    const t = String(row[i] || '')
      .replace(/\s+/g, ' ')
      .trim();
    if (isFormDGJGujaratOuterFootnoteText(t)) {
      const norm = normalizeFormDGJFootnoteCompare(t);
      if (/^\*\*\s*not\s+necessary/.test(norm)) return FORM_DGJ_GJ_ELECTRONIC_FORMAT_NOTE_2;
      if (/^\*\s*not\s+necessary/.test(norm)) return FORM_DGJ_GJ_ELECTRONIC_FORMAT_NOTE_1;
      if (/governor\s+of\s+gujarat/.test(norm)) return FORM_DGJ_GJ_GOVERNOR_ORDER_NOTE;
      return t;
    }
  }
  return '';
};

const looksLikeFormDGJMusterRollPdfContext = (metaLines, rows, sheetName = '') => {
  const blob = [...(metaLines || []), ...(rows || []).slice(0, 22).flat(), sheetName || '']
    .join(' ')
    .toLowerCase();
  if (/rajasthan|\b_rj\b|form[\s._-]*d[\s._-]*rj|form_d_rj/.test(blob)) return false;
  const hasFormD = /\bform\s*d\b|form[\s._-]*d[\s._-]/.test(blob);
  const hasMuster =
    /attendance\s*\/?\s*muster|muster-?roll\s+register|format\s+of\s+attendance|relay\s+or\s+set\s+work/.test(
      blob
    );
  const hasGujarat = /gujarat|\b_gj\b|form[\s._-]*d[\s._-]*gj|form_d_gj/.test(blob);
  if (hasGujarat && (hasFormD || hasMuster)) return true;
  if (hasFormD && hasMuster) return true;
  return false;
};

const isUnpaidAccumulationsFootnoteText = (text) => {
  const norm = String(text || '')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
  if (!norm) return false;
  if (/see\s+definition\s+of/.test(norm) && /unpaid\s+accumulations/.test(norm)) return true;
  if (
    /under\s+section\s*2\s*\(?\s*i\s*\)?/.test(norm) &&
    /labour\s+welfare\s+fund\s+act/.test(norm)
  ) {
    return true;
  }
  return false;
};

const isFormCWhereverApplicableFootnoteText = (text) => {
  const norm = String(text || '')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
  if (!norm) return false;
  return /^\*?\s*wherever\s+applicable\.?$/.test(norm);
};

const isFormCDamageLossFineFootnoteText = (text) => {
  const norm = String(text || '')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
  if (!norm) return false;
  return /applicable\s+only\s+in\s+case\s+of\s+damage\s*\/?\s*loss\s*\/?\s*fine/.test(norm);
};

const isFormCWhereverApplicableFootnoteRow = (row) => {
  if (!Array.isArray(row)) return false;
  const filled = row.map((c) => String(c || '').trim()).filter(Boolean);
  if (!filled.length) return false;
  const unique = [...new Set(filled)];
  return unique.length === 1 && isFormCWhereverApplicableFootnoteText(unique[0]);
};

const isFormCDamageLossFineFootnoteRow = (row) => {
  if (!Array.isArray(row)) return false;
  const filled = row.map((c) => String(c || '').trim()).filter(Boolean);
  if (!filled.length) return false;
  const unique = [...new Set(filled)];
  return unique.length === 1 && isFormCDamageLossFineFootnoteText(unique[0]);
};

const extractFormCWhereverApplicableFootnoteText = (row) => {
  if (!Array.isArray(row)) return '';
  for (let i = 0; i < row.length; i += 1) {
    const t = String(row[i] || '')
      .replace(/\s+/g, ' ')
      .trim();
    if (isFormCWhereverApplicableFootnoteText(t)) return t.startsWith('*') ? t : `*${t}`;
  }
  return '';
};

const isUnpaidAccumulationsFootnoteRow = (row) => {
  if (!Array.isArray(row)) return false;
  const filled = row.map((c) => String(c || '').trim()).filter(Boolean);
  if (!filled.length) return false;
  const unique = [...new Set(filled)];
  return unique.length === 1 && isUnpaidAccumulationsFootnoteText(unique[0]);
};

const looksLikeFormCRajasthanPdfContext = (metaLines, rows, sheetName = '') => {
  const blob = [...(metaLines || []), ...(rows || []).slice(0, 18).flat(), sheetName || '']
    .join(' ')
    .toLowerCase();
  // Gujarat Form C (Loan/Recoveries) must not inherit the RJ damage/loss/fine footnote.
  if (/gujarat|\b_gj\b|form[\s._-]*c[\s._-]*gj|form_c_gj/.test(blob)) return false;
  if (
    /register\s+of\s+(?:loan|recoveries|loan\s*\/\s*recoveries)/.test(blob) &&
    !/rajasthan|\b_rj\b|form[\s._-]*c[\s._-]*rj|form_c_rj/.test(blob)
  ) {
    return false;
  }
  return (
    /\bform\s*c\b|form[\s._-]*c[\s._-]/.test(blob) &&
    (/rajasthan|\b_rj\b|form[\s._-]*c[\s._-]*rj|form_c_rj/.test(blob) ||
      /register\s+of\s+deductions/.test(blob)) &&
    /damage|loss|fine|deduction/.test(blob)
  );
};

const extractUnpaidAccumulationsFootnoteText = (row) => {
  if (!Array.isArray(row)) return '';
  for (let i = 0; i < row.length; i += 1) {
    const t = String(row[i] || '')
      .replace(/\s+/g, ' ')
      .trim();
    if (isUnpaidAccumulationsFootnoteText(t)) return t;
  }
  return '';
};

/** Form 14 RJ overtime footnote — paint below the grid, above the system-generated note. */
const isForm14RajasthanOvertimeFootnoteText = (text) => {
  const norm = String(text || '')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
  if (!norm) return false;
  if (/need\s+not\s+be\s+filled/.test(norm) && /commercial\s+establishments?/.test(norm)) {
    return true;
  }
  return (
    /red\s+ink/.test(norm) &&
    /overtime/.test(norm) &&
    (/days\s+column|commercial\s+establishments?|public\s+amusement/.test(norm) ||
      /^\*\s*this\s+column/.test(norm))
  );
};

const isForm14RajasthanDayEntriesNoteText = (text) => {
  const norm = String(text || '')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
  if (!norm) return false;
  return (
    /note\s*:\s*entries\s+relating\s+to\s+any\s+day/.test(norm) ||
    (/entries\s+relating\s+to\s+any\s+day/.test(norm) && /made\s+on\s+that\s+day/.test(norm))
  );
};

const isForm14RajasthanFootnoteText = (text) =>
  isForm14RajasthanOvertimeFootnoteText(text) || isForm14RajasthanDayEntriesNoteText(text);

const isForm14RajasthanFootnoteRow = (row) => {
  if (!Array.isArray(row)) return false;
  const filled = row.map((c) => String(c || '').trim()).filter(Boolean);
  if (!filled.length) return false;
  const unique = [...new Set(filled)];
  // Merged footnote often collapses to one cell; fragments may repeat across the band.
  return unique.every((t) => isForm14RajasthanFootnoteText(t));
};

const extractForm14RajasthanFootnoteText = (row) => {
  if (!Array.isArray(row)) return '';
  const parts = [];
  for (let i = 0; i < row.length; i += 1) {
    const t = String(row[i] || '')
      .replace(/\s+/g, ' ')
      .trim();
    if (!t || !isForm14RajasthanFootnoteText(t)) continue;
    if (!parts.includes(t)) parts.push(t);
  }
  return parts.join(' ').replace(/\s+/g, ' ').trim();
};

const normalizeForm14RajasthanPdfFootnoteText = (text) => {
  const raw = String(text || '')
    .replace(/\s+/g, ' ')
    .trim();
  if (!raw) return '';
  if (isForm14RajasthanDayEntriesNoteText(raw)) return FORM_14_RJ_DAY_ENTRIES_NOTE;
  if (isForm14RajasthanOvertimeFootnoteText(raw)) {
    // Prefer the captured sheet wording when it is already a full sentence.
    if (raw.length > 80) return raw;
    return FORM_14_RJ_FOOTER_NOTE;
  }
  return raw;
};

const isZipArrayBuffer = (buf) => {
  if (!buf || buf.byteLength < 4) return false;
  const u8 = buf instanceof Uint8Array ? buf : new Uint8Array(buf);
  return u8[0] === 0x50 && u8[1] === 0x4b;
};

const cellToText = (value) => {
  if (value == null || value === '') return '';
  if (typeof value === 'string' || typeof value === 'number' || typeof value === 'boolean') {
    return String(value).trim();
  }
  if (value instanceof Date) {
    try {
      return value.toLocaleDateString('en-IN');
    } catch (_) {
      return value.toISOString().slice(0, 10);
    }
  }
  if (typeof value === 'object') {
    if (value.richText && Array.isArray(value.richText)) {
      return value.richText
        .map((p) => p?.text || '')
        .join('')
        .trim();
    }
    if (value.t === 'd' && value.v instanceof Date) return cellToText(value.v);
    if (value.w != null && String(value.w).trim() !== '') return String(value.w).trim();
    if (value.text != null) return String(value.text).trim();
    if (value.result != null) return cellToText(value.result);
    if (value.v != null) return cellToText(value.v);
    if (value.hyperlink != null && value.text != null) return String(value.text).trim();
  }
  const text = String(value).trim();
  return text === '[object Object]' ? '' : text;
};

const collectExcelBuffersFromDraft = async (arrayBuffer, fileName = '') => {
  const nameHint = String(fileName || '');
  if (!arrayBuffer || arrayBuffer.byteLength < 32) return [];

  const asSingleWorkbook = () => [
    {
      arrayBuffer,
      label: nameHint.replace(EXCEL_EXT_RE, '').replace(/\.zip$/i, '') || 'Draft'
    }
  ];

  if (EXCEL_EXT_RE.test(nameHint) && !/\.zip$/i.test(nameHint)) {
    return asSingleWorkbook();
  }

  if (isZipArrayBuffer(arrayBuffer) || /\.zip$/i.test(nameHint)) {
    try {
      const zip = await JSZip.loadAsync(arrayBuffer);
      const paths = Object.keys(zip.files || {});
      const isOoxmlWorkbook = paths.some(
        (p) =>
          /(^|\/)\[Content_Types\]\.xml$/i.test(p) ||
          /(^|\/)xl\/workbook\.xml$/i.test(p) ||
          /(^|\/)xl\/worksheets\//i.test(p)
      );
      if (isOoxmlWorkbook) return asSingleWorkbook();

      const entries = paths
        .filter((path) => {
          const f = zip.files[path];
          return f && !f.dir && EXCEL_EXT_RE.test(path) && !path.startsWith('__MACOSX');
        })
        .sort((a, b) => a.localeCompare(b))
        .slice(0, MAX_ZIP_EXCEL_FILES);

      if (entries.length > 0) {
        const out = [];
        for (const path of entries) {
          const buf = await zip.files[path].async('arraybuffer');
          if (!buf || buf.byteLength < 32) continue;
          const base = path.split('/').pop() || path;
          out.push({ arrayBuffer: buf, label: base.replace(EXCEL_EXT_RE, '') });
        }
        if (out.length > 0) return out;
      }
    } catch (zipErr) {
      console.warn('Draft ZIP inspect failed; trying as workbook:', zipErr);
    }
  }

  return asSingleWorkbook();
};

/**
 * Keep merged-cell value only in the top-left cell; clear duplicates in the merge range.
 * Prevents FORM XVIII title/address repeating in every column.
 */
const collapseMergedCellDuplicates = (rows, merges) => {
  if (!Array.isArray(merges) || !merges.length || !rows.length) return rows;
  merges.forEach((m) => {
    if (!m?.s || !m?.e) return;
    const r0 = m.s.r;
    const c0 = m.s.c;
    const r1 = m.e.r;
    const c1 = m.e.c;
    let master = '';
    if (rows[r0] && rows[r0][c0] != null) master = cellToText(rows[r0][c0]);
    // If master empty, take first non-empty in range
    if (!master) {
      for (let r = r0; r <= r1 && r < rows.length; r += 1) {
        for (let c = c0; c <= c1; c += 1) {
          const t = cellToText(rows[r]?.[c]);
          if (t) {
            master = t;
            break;
          }
        }
        if (master) break;
      }
    }
    for (let r = r0; r <= r1 && r < rows.length; r += 1) {
      if (!rows[r]) continue;
      while (rows[r].length <= c1) rows[r].push('');
      for (let c = c0; c <= c1; c += 1) {
        rows[r][c] = r === r0 && c === c0 ? master : '';
      }
    }
  });
  return rows;
};

/**
 * If the same long text is copied across many columns (broken merge), keep it once.
 */
const dedupeRepeatedRowText = (row) => {
  if (!Array.isArray(row) || row.length < 3) return row;
  const nonEmpty = row
    .map((c, idx) => ({ text: cellToText(c), idx }))
    .filter((x) => x.text);
  if (nonEmpty.length < 3) return row;

  const counts = new Map();
  nonEmpty.forEach(({ text }) => {
    counts.set(text, (counts.get(text) || 0) + 1);
  });

  let dominant = '';
  let dominantCount = 0;
  counts.forEach((count, text) => {
    if (count > dominantCount) {
      dominant = text;
      dominantCount = count;
    }
  });

  // Same title/address pasted into most columns
  const ratio = dominantCount / nonEmpty.length;
  if (dominant && dominant.length >= 12 && dominantCount >= 3 && ratio >= 0.5) {
    const out = row.map(() => '');
    out[0] = dominant;
    // Preserve other unique short values (rare)
    nonEmpty.forEach(({ text, idx }) => {
      if (text !== dominant) out[idx] = text;
    });
    return out;
  }
  return row;
};

const isFormMetaText = (text) => {
  const t = String(text || '').toLowerCase();
  if (!t) return false;
  return (
    /^form\s*(xviii|xvii|xix|xv|xiv|[a-z0-9-]+)\b/.test(t) ||
    /see rule|register of wages|muster roll|name and address of contractor|nature and location of work|address of the establishment|principal employer|wage period|from\s*to|list of workers|shift schedule|shift scehdule|all the workers in establishment/.test(
      t
    )
  );
};

/** Match Sr.No / S.No / SI.No and common wage-register identity headers. */
const isWageRegisterColHeaderBlob = (blob) => {
  const t = String(blob || '').toLowerCase();
  if (!t) return false;
  return (
    /(?:^|[^a-z])s\.?\s*no(?:[^a-z]|$)|(?:^|[^a-z])sr\.?\s*no(?:[^a-z]|$)|(?:^|[^a-z])si\.?\s*no(?:[^a-z]|$)/.test(
      t
    ) ||
    /name of(?:\s+the)?\s+employee|employee identification|name of the work(?:er|man)|designation|daily (?:attendance|hours)|amount of wages|net amount|rate of wages|basic wage|number of days worked/.test(
      t
    )
  );
};

/**
 * Tamil Nadu LWF Form C — "Details of Fines…" + "Quarter ending…" column headers.
 * Must stay in the grid (never promoted to full-width meta bands).
 */
const isFormCLwfColHeaderBlob = (blob) => {
  const t = String(blob || '').toLowerCase();
  if (!t) return false;
  const quarterHits = t.match(/quarter\s+ending/g);
  if (quarterHits && quarterHits.length >= 2) return true;
  return /details\s+of\s+fines/.test(t) && /quarter\s+ending/.test(t);
};

/**
 * Tamil Nadu Form 25 — S.No / Name / day-band column headers.
 * Must stay in the grid (not promoted to full-width meta).
 */
const isForm25TamilNaduColHeaderBlob = (blob) => {
  const t = String(blob || '').toLowerCase();
  if (!t) return false;
  if (
    /daily\s+hours\s+of\s+work/.test(t) &&
    (/s\.?\s*no|name of the worker|scheme of shifts/.test(t) || /\b1\b.*\b2\b.*\b3\b/.test(t))
  ) {
    return true;
  }
  if (
    /name of the worker/.test(t) &&
    /worker\s+identit|time at which work|scheme of shifts|rest\s+interval/.test(t)
  ) {
    return true;
  }
  return false;
};

/**
 * Rajasthan Form 14 — Record of Hours of Work column headers.
 * Long labels must stay in the grid (never full-width meta bands).
 */
const isForm14RajasthanColHeaderBlob = (blob) => {
  const t = String(blob || '').toLowerCase();
  if (!t) return false;
  if (/name\s+of\s+persons?\s+employ/.test(t) && /young\s+person|total\s+hours\s+worked|overtime/.test(t)) {
    return true;
  }
  if (
    /whether\s+young\s+person/.test(t) &&
    /total\s+hours\s+worked|days\s+on\s+which\s+overtime|extent\s+of\s+overtime/.test(t)
  ) {
    return true;
  }
  return false;
};

/** CLRA Form XIX — Wage Slip (Rule 78) — compact Excel layout must stay portrait in PDF. */
const looksLikeFormXIXWageSlipPdfContext = (metaLines, rows, sheetName = '') => {
  const blob = [...(metaLines || []), ...(rows || []).slice(0, 16).flat(), sheetName || '']
    .join(' ')
    .toLowerCase();
  if (!blob) return false;
  const isXix =
    /form\s*xix\b/.test(blob) ||
    /form[\s._-]*xix[\s._-]/.test(blob) ||
    (/wage\s*s?\s*slip/.test(blob) && /rule\s*78/.test(blob));
  if (!isXix) return false;
  return (
    /no\.?\s*of\s+days\s+worked|days\s+worked/.test(blob) &&
    (/overtime\s+hours|dates\s+on\s+which\s+overtime|gross\s+wages\s+payable|actual\s+wages\s+paid/.test(
      blob
    ) ||
      /rate\s+of\s+daily\s+wages/.test(blob))
  );
};

const isFormXIXWageSlipColHeaderBlob = (blob) => {
  const t = String(blob || '').toLowerCase();
  if (!t) return false;
  return (
    /no\.?\s*of\s+days\s+worked/.test(t) &&
    (/rate\s+of\s+daily\s+wages|units\s+worked|overtime/.test(t) ||
      /dates\s+on\s+which\s+overtime/.test(t))
  );
};

const isFormXIXAPWageSlipPdfContext = (metaLines, rows, sheetName = '') => {
  const blob = [...(metaLines || []), ...(rows || []).slice(0, 24).flat(), sheetName || '']
    .join(' ')
    .toLowerCase();
  // Gujarat / Karnataka Form XIX have their own layouts — do not treat as AP.
  if (looksLikeFormXIXGJWageSlipPdfContext(metaLines, rows, sheetName)) return false;
  if (looksLikeFormXIXKarnatakaPdfContext(metaLines, rows, sheetName)) return false;
  if (/gujarat|form[\s._-]*xix[\s._-]*gj|xix_gj|\bgj[-_\s]/.test(blob)) return false;
  if (/karnataka|form[\s._-]*xix[\s._-]*ka\b|xix_ka/.test(blob)) return false;
  return (
    /form[\s._-]*xix\b/.test(blob) &&
    (/andhra\s+pradesh|\bap\b|form[\s._-]*xix[\s._-]*ap\b/.test(blob) ||
      (/nature\s+and\s+location\s+of\s+work/.test(blob) &&
        /gross\s+wages\s+payable/.test(blob) &&
        /net\s+amount\s+of\s+wages/.test(blob)))
  );
};

const FORM_XIX_PDF_WAGE_HEADERS = [
  'No. of days worked',
  'Rate of daily wages/piece - rate',
  'No. of units worked in case of piece rate',
  'Dates on which overtime worked',
  'Overtime hours and amount of overtime wages'
];

const FORM_XIX_PDF_FOOTER_LABELS = [
  'Gross wages payable',
  'Deductions, any',
  'If Actual wages paid'
];

const pickFormXIXPdfNonEmptyCells = (row) =>
  (Array.isArray(row) ? row : [])
    .map((c) => String(c || '').trim())
    .filter(Boolean);

/**
 * Collapse sparse Excel merges (A–M) into the official 5 wage columns + footer band
 * so Download PDF matches the Excel Form XIX wage-slip layout.
 */
const normalizeFormXIXWageSlipPdfMatrix = (rows, colCount, tableStartRow = 0) => {
  const src = Array.isArray(rows) ? rows : [];
  if (!src.length) return { rows: src, colCount, tableStartRow };

  let headerRow = -1;
  for (let r = Math.max(0, tableStartRow - 2); r < Math.min(src.length, tableStartRow + 8); r += 1) {
    const blob = pickFormXIXPdfNonEmptyCells(src[r]).join(' ').toLowerCase();
    if (isFormXIXWageSlipColHeaderBlob(blob) || /no\.?\s*of\s+days\s+worked/.test(blob)) {
      headerRow = r;
      break;
    }
  }
  if (headerRow < 0) {
    for (let r = 0; r < Math.min(src.length, 24); r += 1) {
      const blob = pickFormXIXPdfNonEmptyCells(src[r]).join(' ').toLowerCase();
      if (isFormXIXWageSlipColHeaderBlob(blob)) {
        headerRow = r;
        break;
      }
    }
  }
  if (headerRow < 0) return { rows: src, colCount, tableStartRow };

  // Keep original column indexes so blank "units" does not shift OT / NIL values left.
  const headerSrc = Array.isArray(src[headerRow]) ? src[headerRow] : [];
  const headerColIndexes = [];
  for (let c = 0; c < headerSrc.length; c += 1) {
    if (String(headerSrc[c] || '').trim()) headerColIndexes.push(c);
  }
  while (headerColIndexes.length < 5) {
    const last = headerColIndexes.length ? headerColIndexes[headerColIndexes.length - 1] + 1 : 0;
    headerColIndexes.push(last);
  }
  const wageCols = headerColIndexes.slice(0, 5);

  const cellAt = (rowIndex, colIndex) =>
    String(src[rowIndex]?.[colIndex] ?? '').trim();

  const headers = FORM_XIX_PDF_WAGE_HEADERS.map(
    (fallback, i) => cellAt(headerRow, wageCols[i]) || fallback
  );

  let numberRow = -1;
  for (let r = headerRow + 1; r <= Math.min(src.length - 1, headerRow + 3); r += 1) {
    const filled = wageCols.map((c) => cellAt(r, c)).filter(Boolean);
    if (filled.length >= 3 && filled.every((t) => /^\d{1,2}$/.test(t))) {
      numberRow = r;
      break;
    }
  }

  let dataRow = -1;
  for (
    let r = (numberRow >= 0 ? numberRow : headerRow) + 1;
    r <= Math.min(src.length - 1, headerRow + 6);
    r += 1
  ) {
    const filled = wageCols.map((c) => cellAt(r, c));
    if (!filled.some(Boolean)) continue;
    if (filled.filter(Boolean).every((t) => /^\d{1,2}$/.test(t))) continue;
    const blob = filled.join(' ').toLowerCase();
    if (/gross\s+wages|deductions|actual\s+wages|signature/.test(blob)) continue;
    dataRow = r;
    break;
  }

  let footerLabelRow = -1;
  let footerValueRow = -1;
  for (let r = Math.max(headerRow + 1, dataRow + 1); r < Math.min(src.length, headerRow + 14); r += 1) {
    const blob = pickFormXIXPdfNonEmptyCells(src[r]).join(' ').toLowerCase();
    if (/gross\s+wages\s+payable/.test(blob) && /deduction|actual\s+wages/.test(blob)) {
      footerLabelRow = r;
      break;
    }
    if (/gross\s+wages\s+payable/.test(blob)) {
      footerLabelRow = r;
      break;
    }
  }
  if (footerLabelRow >= 0) {
    for (let r = footerLabelRow + 1; r <= Math.min(src.length - 1, footerLabelRow + 2); r += 1) {
      const filled = pickFormXIXPdfNonEmptyCells(src[r]);
      if (!filled.length) continue;
      const blob = filled.join(' ').toLowerCase();
      if (/signature|contractor|representative/.test(blob)) continue;
      if (filled.some((t) => /^-?\d+(\.\d+)?$/.test(String(t).replace(/,/g, '')))) {
        footerValueRow = r;
        break;
      }
    }
  }

  const dataValues = wageCols.map((c) => (dataRow >= 0 ? cellAt(dataRow, c) : ''));

  // Footer labels/values sit under wage cols 1 / 3 / 5 in the Excel template.
  const footerAnchorCols = [wageCols[0], wageCols[2], wageCols[4]];
  const footerLabels = FORM_XIX_PDF_FOOTER_LABELS.map((fallback, i) => {
    if (footerLabelRow < 0) return fallback;
    const fromAnchor = cellAt(footerLabelRow, footerAnchorCols[i]);
    if (fromAnchor) return fromAnchor;
    const packed = pickFormXIXPdfNonEmptyCells(src[footerLabelRow]);
    return packed[i] || fallback;
  });

  const footerValues = ['', '', '', '', ''];
  if (footerValueRow >= 0) {
    const fromAnchors = footerAnchorCols.map((c) => cellAt(footerValueRow, c));
    if (fromAnchors.some(Boolean)) {
      footerValues[0] = fromAnchors[0];
      footerValues[2] = fromAnchors[1];
      footerValues[4] = fromAnchors[2];
    } else {
      const packed = pickFormXIXPdfNonEmptyCells(src[footerValueRow]);
      footerValues[0] = packed[0] || '';
      footerValues[2] = packed[1] || '';
      footerValues[4] = packed[2] || '';
    }
  }

  const compact = [];
  for (let r = 0; r < headerRow; r += 1) {
    compact.push(['', '', '', '', '']);
  }
  compact.push(headers);
  compact.push(['1', '2', '3', '4', '5']);
  compact.push(dataValues);
  compact.push(['', '', '', '', '']);
  compact.push([footerLabels[0], '', footerLabels[1], '', footerLabels[2]]);
  compact.push(footerValues);

  const afterFooter = Math.max(footerValueRow, footerLabelRow, dataRow, headerRow) + 1;
  for (let r = afterFooter; r < src.length; r += 1) {
    const filled = pickFormXIXPdfNonEmptyCells(src[r]);
    if (!filled.length) continue;
    compact.push([filled.join(' '), '', '', '', '']);
  }

  return {
    rows: compact,
    colCount: 5,
    tableStartRow: headerRow
  };
};

const applyFormXIXKarnatakaPdfNormalization = (target, normalized) => {
  if (!target || !normalized) return target;
  target.rows = normalized.rows;
  target.colCount = normalized.colCount;
  target.tableStartRow = normalized.tableStartRow;
  if (Array.isArray(normalized.metaLines)) target.metaLines = normalized.metaLines;
  target.formXIXKarnatakaLayout = normalized.formXIXKarnatakaLayout === true;
  target.formXIXKABottomSection = normalized.formXIXKABottomSection || null;
  return target;
};

const normalizeFormXIXWageSlipPdfByState = (
  rows,
  colCount,
  tableStartRow,
  metaLines = [],
  sheetName = '',
  fileName = ''
) => {
  if (looksLikeFormXIXKarnatakaPdfContext(metaLines, rows, sheetName, fileName)) {
    return normalizeFormXIXKarnatakaWageSlipPdfMatrix(rows, colCount, tableStartRow, metaLines);
  }
  if (looksLikeFormXIXGJWageSlipPdfContext(metaLines, rows, sheetName, fileName)) {
    return normalizeFormXIXGJWageSlipPdfMatrix(rows, colCount, tableStartRow, metaLines);
  }
  if (looksLikeFormXIXMPCLRAPdfContext(metaLines, rows, sheetName)) {
    return normalizeFormXIXMPCLRAWageSlipPdfMatrix(rows, colCount, tableStartRow, metaLines);
  }
  if (isFormXIXAPWageSlipPdfContext(metaLines, rows, sheetName)) {
    return normalizeFormXIXAPWageSlipPdfMatrix(rows, colCount, tableStartRow, metaLines);
  }
  return normalizeFormXIXWageSlipPdfMatrix(rows, colCount, tableStartRow);
};

const FORM_XIX_AP_WAGE_LABEL_RE =
  /^(?:[1-7][.)]?\s*)?(?:no\.?\s+of\s+days\s+worked|no\.?\s+of\s+units\s+worked|rate\s+of\s+daily\s+wages|amount\s+of\s+overtime\s+wages|gross\s+wages\s+payable|deductions?,?\s+if\s+any|net\s+amount\s+of\s+wages)/i;

const isFormXIXAPWageLabel = (text) => FORM_XIX_AP_WAGE_LABEL_RE.test(String(text || '').trim());
const isFormXIXAPMoneyLabel = (text) =>
  /amount\s+of\s+overtime\s+wages|gross\s+wages\s+payable|deductions?,?\s+if\s+any|net\s+amount\s+of\s+wages/i.test(
    String(text || '')
  );

const formatFormXIXAPMoneyValue = (label, value) => {
  const raw = String(value || '').trim();
  if (!raw || !isFormXIXAPMoneyLabel(label)) return raw;
  const normalized = raw.replace(/,/g, '');
  if (!/^-?\d+(?:\.\d+)?$/.test(normalized)) return raw;
  const amount = Number(normalized);
  return Number.isFinite(amount)
    ? amount.toLocaleString('en-IN', { maximumFractionDigits: 2 })
    : raw;
};

const formXIXAPHeaderFieldLabel = (text) => {
  const value = String(text || '').replace(/\s+/g, ' ').trim();
  if (!value) return '';
  return /^(?:name\s+and\s+address\s+of\s+contractor|nature\s+and\s+location\s+of\s+work|name\s+and\s+father|father.*husband.*name|for\s+the\s+(?:week|fortnight|month))/i.test(
    value
  )
    ? value
    : '';
};

const normalizeFormXIXAPWageSlipPdfMatrix = (rows, colCount, tableStartRow = 0, metaLines = []) => {
  const src = Array.isArray(rows) ? rows : [];
  const titles = [];
  const headerPairs = [];
  const wagePairs = [];
  const trailingPairs = [];
  const footerLines = [];
  const consumed = new Set();
  const addPair = (target, label, value = '') => {
    const left = String(label || '').replace(/\s+/g, ' ').trim();
    const right = String(value || '').trim();
    if (!left && !right) return;
    target.push([left, right]);
  };

  for (let metaIndex = 0; metaIndex < (metaLines || []).length; metaIndex += 1) {
    const line = metaLines[metaIndex];
    const text = String(line || '').replace(/\s+/g, ' ').trim();
    if (!text) continue;
    if (/initials\s+of\s+the\s+contractor|signature\s+of\s+the\s+contractor/i.test(text)) {
      footerLines.push(text);
      continue;
    }
    const colon = text.indexOf(':');
    const dotted = text.match(
      /^(name\s+and\s+address\s+of\s+contractor|nature\s+and\s+location\s+of\s+work|name\s+and\s+father[^:.-]*workman|for\s+the\s+(?:week|fortnight|month)[^:.-]*)(?:\s*:\s*|\s*\.{2,}\s*|\s+-\s*)(.*)$/i
    );
    const label = formXIXAPHeaderFieldLabel(
      dotted?.[1] || (colon >= 0 ? text.slice(0, colon) : text)
    );
    if (label) {
      let value = dotted ? dotted[2] : colon >= 0 ? text.slice(colon + 1) : '';
      const next = String(metaLines[metaIndex + 1] || '').replace(/\s+/g, ' ').trim();
      if (!value && next && !formXIXAPHeaderFieldLabel(next) && !/^form\s*xix\b|^wage\s+slip\b|^\[?\s*rule\s*78/i.test(next)) {
        value = next;
        metaIndex += 1;
      }
      addPair(headerPairs, label, value);
    } else if (/^form\s*xix\b|^wage\s+slip\b|^\[?\s*rule\s*78/i.test(text)) {
      titles.push(text);
    } else {
      addPair(headerPairs, text, '');
    }
  }

  const wageStart = Math.max(0, tableStartRow);
  for (let r = wageStart; r < src.length; r += 1) {
    const row = src[r] || [];
    const labelIndex = row.findIndex((cell) => isFormXIXAPWageLabel(cell));
    if (labelIndex < 0) continue;
    const label = String(row[labelIndex] || '').replace(/\s+/g, ' ').trim();
    const sameRowValues = row
      .slice(labelIndex + 1)
      .map((cell) => String(cell || '').trim())
      .filter((cell) => cell && !/^\(?\s*\d{1,2}\s*\)?$/.test(cell));
    let value = sameRowValues.join(' ');
    if (!value) {
      const nextRow = src[r + 1] || [];
      const nextFilled = nextRow.map((cell) => String(cell || '').trim()).filter(Boolean);
      if (nextFilled.length && !nextFilled.some(isFormXIXAPWageLabel)) {
        value = nextFilled.join(' ');
        consumed.add(r + 1);
      }
    }
    addPair(wagePairs, label, value);
    consumed.add(r);
  }

  for (let r = 0; r < src.length; r += 1) {
    if (consumed.has(r)) continue;
    if (r < tableStartRow) continue;
    const filled = (src[r] || []).map((cell) => String(cell || '').trim()).filter(Boolean);
    if (!filled.length) continue;
    if (isSystemGeneratedDocumentNoteRow(src[r])) {
      addPair(trailingPairs, filled.join(' '), '');
      continue;
    }
    const joined = filled.join(' ');
    if (/initials\s+of\s+the\s+contractor|signature\s+of\s+the\s+contractor/i.test(joined)) {
      footerLines.push(joined);
      continue;
    }
    if (filled.some(isFormXIXAPWageLabel)) {
      const labelIndex = filled.findIndex(isFormXIXAPWageLabel);
      addPair(wagePairs, filled[labelIndex], filled.slice(labelIndex + 1).join(' '));
      continue;
    }
    if (filled.length > 1) addPair(trailingPairs, filled[0], filled.slice(1).join(' '));
  }

  const workersPair = [...headerPairs, ...trailingPairs].find(
    (pair) => /^workers?$/i.test(pair[0]) && !pair[1]
  );
  if (workersPair) {
    const unitsPair = wagePairs.find((pair) => /no\.?\s+of\s+units\s+worked/i.test(pair[0]));
    if (unitsPair && !/\bworkers?$/i.test(unitsPair[0])) unitsPair[0] = `${unitsPair[0]} ${workersPair[0]}`;
    [headerPairs, trailingPairs].forEach((list) => {
      const index = list.indexOf(workersPair);
      if (index >= 0) list.splice(index, 1);
    });
  }

  const uniquePairs = [];
  [...headerPairs, ...wagePairs, ...trailingPairs].forEach((pair) => {
    if (!pair[0] && !pair[1]) return;
    if (!uniquePairs.some((existing) => existing[0] === pair[0] && existing[1] === pair[1])) {
      uniquePairs.push(pair);
    }
  });
  return {
    rows: uniquePairs,
    colCount: 2,
    tableStartRow: 0,
    metaLines: titles,
    formXIXAPLayout: true,
    formXIXAPFooterLines: footerLines
  };
};

/** CLRA Form XIV / Employment Card (Rule 76) — stacked label|value card, not a multi-col register. */
const looksLikeFormXIVEmploymentCardPdfContext = (metaLines, rows, sheetName = '') => {
  const blob = [...(metaLines || []), ...(rows || []).slice(0, 24).flat(), sheetName || '']
    .join(' ')
    .toLowerCase();
  if (!blob) return false;
  // Rajasthan Form X Employment Card is a horizontal workman table — keep that layout.
  if (
    (/form[\s._-]*x[\s._-]*rj\b|\bx_rj\b/.test(blob) || /see\s+rule\s*75/.test(blob)) &&
    /employment\s+card/.test(blob) &&
    !/form\s*xiv\b/.test(blob)
  ) {
    return false;
  }
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

const looksLikeFormXRajasthanEmploymentCardPdfContext = (metaLines, rows, sheetName = '') => {
  const blob = [...(metaLines || []), ...(rows || []).slice(0, 24).flat(), sheetName || '']
    .join(' ')
    .toLowerCase();
  return (
    /see\s+rule\s*75/.test(blob) &&
    /employment\s+card/.test(blob) &&
    /name\s+and\s+address\s+of\s+(?:the\s+)?contractor/.test(blob) &&
    !/form\s*xiv\b/.test(blob)
  );
};

const looksLikeFormXIRajasthanServiceCertificatePdfContext = (metaLines, rows, sheetName = '') => {
  const blob = [...(metaLines || []), ...(rows || []).slice(0, 28).flat(), sheetName || '']
    .join(' ')
    .toLowerCase();
  return (
    /see\s+rule\s*76/.test(blob) &&
    /service\s+certificate/.test(blob) &&
    /form\s*xi\b|form[\s._-]*xi[\s._-]/.test(blob)
  );
};

const FORM_XI_RJ_PDF_HEADER_LABELS = [
  'Name and address of contractor',
  'Nature and location of work',
  'Name and address of establishment under which contract is carried on',
  'Name and address of principal employer',
  'Name and address of the workman',
  'Age or date of birth',
  'Identification marks',
  "Father's / Husband's Name",
];

const matchFormXIRajasthanServiceHeaderIndex = (text) => {
  const normalized = String(text || '')
    .replace(/\.+$/g, '')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
  if (!normalized) return -1;
  if (/name\s+and\s+address\s+(?:of|if)\s+(?:the\s+)?contractor/.test(normalized)) return 0;
  if (/nature\s+and\s+location\s+of\s+work/.test(normalized)) return 1;
  if (/name\s+and\s+address\s+of\s+(?:the\s+)?establishment/.test(normalized)) return 2;
  if (/name\s+and\s+address\s+of\s+(?:the\s+)?principal\s+employer/.test(normalized)) return 3;
  if (/name\s+and\s+address\s+of\s+the\s+workm[ae]n/.test(normalized)) return 4;
  if (/age\s+or\s+date\s+of\s+birth/.test(normalized)) return 5;
  if (/identification\s+marks/.test(normalized)) return 6;
  if (/father.*husband.*name|husband.*father.*name/.test(normalized)) return 7;
  return -1;
};

const buildFormXIRajasthanServiceHeaderModel = (metaLines) => {
  const lines = [];
  (metaLines || []).forEach((raw) => {
    expandStatutoryMetaSegments(raw).forEach((line) => {
      const text = String(line || '').replace(/\s+/g, ' ').trim();
      if (text) lines.push(text);
    });
  });
  const values = FORM_XI_RJ_PDF_HEADER_LABELS.map(() => '');
  const titles = [];
  let activeIndex = -1;
  lines.forEach((line) => {
    const index = matchFormXIRajasthanServiceHeaderIndex(line);
    if (index >= 0) {
      activeIndex = index;
      const inline = line.replace(/^[^:]*:\s*/, '').trim();
      if (inline && inline !== line) values[index] = inline;
      return;
    }
    if (activeIndex >= 0 && !values[activeIndex]) {
      if (!/^form\s+xi\b|^service\s+certificate$|^see\s+rule/i.test(line)) {
        values[activeIndex] = line;
        activeIndex = -1;
        return;
      }
    }
    if (activeIndex < 0 && !/^sample$|^rj[-\s]/i.test(line)) titles.push(line);
  });
  return {
    titles,
    fields: FORM_XI_RJ_PDF_HEADER_LABELS.map((label, index) =>
      `${label}: ${values[index]}`.trimEnd()
    ),
  };
};

const looksLikeFormXVRajasthanWageSlipPdfContext = (metaLines, rows, sheetName = '') => {
  const blob = [...(metaLines || []), ...(rows || []).slice(0, 24).flat(), sheetName || '']
    .join(' ')
    .toLowerCase();
  return (
    /form\s*xv\b|form[\s._-]*xv[\s._-]/.test(blob) &&
    /wage\s*slip|wages\s+slip/.test(blob) &&
    (/see\s+rule\s*77|rajasthan|rule\s*77\s*\(\s*2\s*\)\s*\(\s*b\s*\)/.test(blob)) &&
    !/form\s*xi\b|service\s+certificate/.test(blob)
  );
};

const looksLikeFormXVAPServiceCertificatePdfContext = (metaLines, rows, sheetName = '') => {
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

const FORM_XV_AP_SERVICE_FIELD_LABELS = [
  'Name and address of the workman',
  'Age or date of Birth',
  'Identification Marks',
  "Father's / Husband's Name"
];

const formXVAPServiceFieldIndex = (text) => {
  const t = String(text || '').replace(/\s+/g, ' ').trim().toLowerCase();
  if (/name\s+and\s+address\s+of\s+the\s+workman/.test(t)) return 0;
  if (/age\s+or\s+date\s+of\s+birth/.test(t)) return 1;
  if (/identification\s+marks?/.test(t)) return 2;
  if (/father.*husband.*name|husband.*father.*name/.test(t)) return 3;
  return -1;
};

const buildFormXVAPServiceCertificateHeaderModel = (metaLines, rows = [], tableStartRow = 0) => {
  const lines = [];
  [...(metaLines || []), ...rows.slice(0, Math.max(0, tableStartRow)).flat()].forEach((raw) => {
    expandStatutoryMetaSegments(raw).forEach((line) => {
      const text = String(line || '').replace(/\s+/g, ' ').trim();
      if (text) lines.push(text);
    });
  });
  const titleLines = [];
  const admin = [ ['', ''], ['', ''] ];
  const fields = FORM_XV_AP_SERVICE_FIELD_LABELS.map(() => '');
  let activeField = -1;
  let activeAdmin = null;
  lines.forEach((line) => {
    const lower = line.toLowerCase();
    const fieldIndex = formXVAPServiceFieldIndex(line);
    if (fieldIndex >= 0) {
      activeField = fieldIndex;
      const inline = line.replace(/^[^:]*:\s*/, '').trim();
      if (inline && inline !== line) fields[fieldIndex] = inline;
      activeAdmin = null;
      return;
    }
    if (/name\s+and\s+address\s+of\s+(?:the\s+)?contractor/.test(lower)) {
      activeAdmin = [0, 0];
      const inline = line.replace(/^[^:]*:\s*/, '').trim();
      if (inline && inline !== line) admin[0][0] = inline;
      activeField = -1;
      return;
    }
    if (/name\s+and\s+address\s+of\s+(?:the\s+)?establishment/.test(lower)) {
      activeAdmin = [0, 1];
      const inline = line.replace(/^[^:]*:\s*/, '').trim();
      if (inline && inline !== line) admin[0][1] = inline;
      activeField = -1;
      return;
    }
    if (/nature\s+and\s+location\s+of\s+work/.test(lower)) {
      activeAdmin = [1, 0];
      const inline = line.replace(/^[^:]*:\s*/, '').trim();
      if (inline && inline !== line) admin[1][0] = inline;
      activeField = -1;
      return;
    }
    if (/name\s+and\s+address\s+of\s+(?:the\s+)?principal\s+employer/.test(lower)) {
      activeAdmin = [1, 1];
      const inline = line.replace(/^[^:]*:\s*/, '').trim();
      if (inline && inline !== line) admin[1][1] = inline;
      activeField = -1;
      return;
    }
    if (activeField >= 0 && !fields[activeField] && !/^form\s+xv\b|^service\s+certificate$|^\[?vide\b/i.test(line)) {
      fields[activeField] = line;
      activeField = -1;
      return;
    }
    if (activeAdmin && !admin[activeAdmin[0]][activeAdmin[1]] && !/^form\s+xv\b|^service\s+certificate$|^\[?vide\b/i.test(line)) {
      admin[activeAdmin[0]][activeAdmin[1]] = line;
      activeAdmin = null;
      return;
    }
    if (/^form\s+xv\b|^service\s+certificate$|^\[?vide\b/i.test(line)) titleLines.push(line);
  });
  const seenTitles = new Set();
  const uniqueTitleLines = titleLines.filter((line) => {
    const key = String(line || '').replace(/\s+/g, ' ').trim().toLowerCase();
    if (!key || seenTitles.has(key)) return false;
    seenTitles.add(key);
    return true;
  });
  return {
    titles: uniqueTitleLines,
    adminRows: admin,
    fields: FORM_XV_AP_SERVICE_FIELD_LABELS.map((label, index) => `${label}: ${fields[index]}`.trimEnd()),
    fieldRows: FORM_XV_AP_SERVICE_FIELD_LABELS.map((label, index) => [label, fields[index]])
  };
};

const FORM_XV_RJ_PDF_CANONICAL_TITLES = ['FORM XV', '[See Rule 77(2)(b)]', 'Wages Slip'];

const FORM_XV_RJ_PDF_HEADER_LABELS = [
  'Name and address of contractor',
  'Name and location of work',
  'Name and address of establishment in/under which contract is carried on',
  'Name and address of principal employer',
  "Name and Father's name of the workman",
  'Sex and identification token/ticket No.',
  'For the week/fortnight/month',
];

const matchFormXVRajasthanWageSlipHeaderIndex = (text) => {
  const normalized = String(text || '')
    .replace(/\.+$/g, '')
    .replace(/[.…_·-]{2,}/gu, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
  if (!normalized) return -1;
  if (/name\s+and\s+address\s+of\s+(?:the\s+)?contractor/.test(normalized)) return 0;
  if (/(?:name|nature)\s+and\s+location\s+of\s+work/.test(normalized)) return 1;
  // Tolerate establiishment / Establishemnt template typos.
  if (
    /name\s+and\s+address\s+of\s+(?:the\s+)?establ(?:[a-z]*ment|ishemnt)/.test(normalized) ||
    (/establ(?:[a-z]*ment|ishemnt)/.test(normalized) &&
      /contract/.test(normalized) &&
      /carried\s+on/.test(normalized))
  ) {
    return 2;
  }
  if (/name\s+and\s+address\s+of\s+(?:the\s+)?principal\s+employer/.test(normalized)) return 3;
  if (/name\s+and\s+father.*workman|father.*name\s+of\s+the\s+workman|name\s+of\s+(?:the\s+)?workman/.test(normalized)) {
    return 4;
  }
  if (/sex\s+and\s+identification|identification\s+token|ticket\s+no|sex\s+and\s+id/.test(normalized)) {
    return 5;
  }
  if (
    /for\s+the\s+(?:week|fortnight)|week\s*\/\s*fortnight\s*\/\s*month|for\s+the\s+month\s+of|month\s+ended/
      .test(normalized)
  ) {
    return 6;
  }
  return -1;
};

const buildFormXVRajasthanWageSlipHeaderModel = (metaLines) => {
  const lines = [];
  (metaLines || []).forEach((raw) => {
    expandStatutoryMetaSegments(raw).forEach((line) => {
      const text = String(line || '').replace(/\s+/g, ' ').trim();
      if (text) lines.push(text);
    });
  });
  const values = FORM_XV_RJ_PDF_HEADER_LABELS.map(() => '');
  let activeIndex = -1;
  const isTitleLine = (line) =>
    /^form\s+xv\b|^wages?\s+slip$|^see\s+rule|^\[?see\s+rule/i.test(line);
  const looksLikeAnotherHeader = (line) => matchFormXVRajasthanWageSlipHeaderIndex(line) >= 0;
  const looksLikeEmployeeToken = (line) =>
    /^VE\d+$/i.test(String(line || '').trim()) && String(line).trim().length < 16;
  const looksLikePersonNameLeftover = (line) => {
    const text = String(line || '').trim();
    if (!text || /,/.test(text) || /\d/.test(text)) return false;
    if (/establishment|address|limited|private|road|nagar|district|state|india|pvt|ltd/i.test(text)) {
      return false;
    }
    return /^[A-Za-z][A-Za-z.'-]{1,30}(?:\s+[A-Za-z][A-Za-z.']{1,30}){0,3}$/.test(text);
  };
  const appendFatherNameToWorkman = (line) => {
    const father = String(line || '').trim();
    if (!father || !values[4] || /\n/.test(values[4])) return false;
    if (!looksLikePersonNameLeftover(father)) return false;
    // Avoid duplicating the same workman name as "father".
    if (father.toLowerCase() === String(values[4]).trim().toLowerCase()) return false;
    values[4] = `${values[4]}\n${father}`;
    return true;
  };

  lines.forEach((line) => {
    const index = matchFormXVRajasthanWageSlipHeaderIndex(line);
    if (index >= 0) {
      activeIndex = index;
      // Inline "Label: value" — ignore dotted underline leftovers as value.
      const inline = line.replace(/^[^:]*:\s*/, '').trim();
      const inlineClean = inline.replace(/[.…_·\-\s]+$/gu, '').trim();
      if (
        inlineClean &&
        inlineClean !== line &&
        !looksLikeAnotherHeader(inlineClean) &&
        !isTitleLine(inlineClean)
      ) {
        values[index] = inlineClean;
      }
      return;
    }
    if (isTitleLine(line)) {
      activeIndex = -1;
      return;
    }
    // Sex/token often comes as "Male" then a following "VE1580" line.
    if (values[5] && looksLikeEmployeeToken(line) && !/\bVE\d+\b/i.test(values[5])) {
      values[5] = `${values[5]} ${line.trim()}`.trim();
      activeIndex = -1;
      return;
    }
    // Workman is Excel multiline: name then father's name on the next line.
    if (
      values[4] &&
      (activeIndex === 4 || activeIndex < 0) &&
      !looksLikeAnotherHeader(line) &&
      !looksLikeEmployeeToken(line) &&
      appendFatherNameToWorkman(line)
    ) {
      activeIndex = -1;
      return;
    }
    if (activeIndex >= 0 && !values[activeIndex]) {
      if (!looksLikeAnotherHeader(line)) {
        // Do not treat short sample name/token leftovers as establishment address.
        if (activeIndex === 2) {
          if (looksLikeEmployeeToken(line) || looksLikePersonNameLeftover(line)) return;
        }
        values[activeIndex] = line;
        // Keep workman open so the following father's-name line can append.
        if (activeIndex === 4) return;
        // Keep sex field open so a following token line can append.
        if (activeIndex === 5 && /^(male|female|other|m|f)\b/i.test(line) && !/\bVE\d+\b/i.test(line)) {
          return;
        }
        activeIndex = -1;
        return;
      }
    }
    // Never promote leftover sample names/tokens into PDF titles — Excel always shows
    // FORM XV / [See Rule 77(2)(b)] / Wages Slip.
  });

  return {
    // Match Excel title band exactly (centered form name block).
    titles: [...FORM_XV_RJ_PDF_CANONICAL_TITLES],
    fields: FORM_XV_RJ_PDF_HEADER_LABELS.map((label, index) => {
      const value = String(values[index] || '').trim();
      if (!value) return `${label}:`;
      // Preserve Excel multiline workman (name + father's name).
      if (index === 4 && value.includes('\n')) {
        const [workmanName, ...fatherParts] = value.split(/\n+/).map((s) => s.trim()).filter(Boolean);
        const father = fatherParts.join(' ').trim();
        return father ? `${label}: ${workmanName}\n${father}` : `${label}: ${workmanName}`;
      }
      return `${label}: ${value}`.trimEnd();
    }),
  };
};

const FORM_X_RJ_EMPLOYMENT_HEADER_LABELS = [
  'Name and address of contractor',
  'Nature and location of work',
  'Name and address of establishment under which contract is carried on',
  'Name and address of principal employer',
];

const matchFormXRajasthanEmploymentHeaderIndex = (text) => {
  const normalized = String(text || '')
    .replace(/\.+$/g, '')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
  if (!normalized) return -1;
  if (/name\s+and\s+address\s+(?:of|if)\s+(?:the\s+)?contractor/.test(normalized)) return 0;
  if (/nature\s+and\s+location\s+of\s+work/.test(normalized)) return 1;
  if (/name\s+and\s+address\s+of\s+(?:the\s+)?establishment/.test(normalized)) return 2;
  if (/name\s+and\s+address\s+of\s+(?:the\s+)?principal\s+employer/.test(normalized)) return 3;
  return -1;
};

const buildFormXRajasthanEmploymentHeaderModel = (metaLines) => {
  const lines = [];
  (metaLines || []).forEach((raw) => {
    expandStatutoryMetaSegments(raw).forEach((line) => {
      const text = String(line || '').replace(/\s+/g, ' ').trim();
      if (text) lines.push(text);
    });
  });
  const titles = [];
  const fields = [];
  const values = FORM_X_RJ_EMPLOYMENT_HEADER_LABELS.map(() => '');
  let activeIndex = -1;
  lines.forEach((line) => {
    const index = matchFormXRajasthanEmploymentHeaderIndex(line);
    if (index >= 0) {
      activeIndex = index;
      const inline = line.replace(/^[^:]*:\s*/, '').trim();
      if (inline && inline !== line) values[index] = inline;
      return;
    }
    if (activeIndex >= 0 && !values[activeIndex] && !/^form\s+x\b|^employment\s+card$|^see\s+rule/i.test(line)) {
      values[activeIndex] = line;
      activeIndex = -1;
      return;
    }
    if (activeIndex < 0 && !/^sample$|^rj[-\s]/i.test(line)) titles.push(line);
  });
  FORM_X_RJ_EMPLOYMENT_HEADER_LABELS.forEach((label, index) => {
    fields.push(`${label}: ${values[index]}`.trimEnd());
  });
  return { titles, fields };
};

/** True when several workman headers sit on one row (Form X_RJ tabular card). */
const isFormXIVEmploymentCardTabularHeaderRow = (row) => {
  const filled = (Array.isArray(row) ? row : [])
    .map((c) => String(c || '').trim())
    .filter(Boolean);
  if (filled.length < 3) return false;
  const blob = filled.join(' ').toLowerCase();
  let hits = 0;
  if (/name\s+of\s+the\s+workman/.test(blob)) hits += 1;
  if (/serial|s\.?\s*no/.test(blob) && /register|workman/.test(blob)) hits += 1;
  if (/nature/.test(blob) && /employ|designat/.test(blob)) hits += 1;
  if (/wage\s+period/.test(blob)) hits += 1;
  if (/tenure|period\s+of\s+employ/.test(blob)) hits += 1;
  if (/wage/.test(blob) && /rate|piece/.test(blob)) hits += 1;
  return hits >= 3;
};

const isFormXIVEmploymentCardWorkmanLabelText = (text) => {
  const t = String(text || '')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase()
    .replace(/^\d+[\.\)]\s*/, '');
  if (!t) return false;
  return (
    /name\s+of\s+the\s+workman/.test(t) ||
    (/serial|s\.?\s*no/.test(t) && /register/.test(t)) ||
    (/nature/.test(t) && /employ|designat/.test(t)) ||
    (/wage/.test(t) && /rate|piece|particular|unit/.test(t)) ||
    /wage\s+period/.test(t) ||
    /tenure\s+of\s+employ|period\s+of\s+employ/.test(t) ||
    /^remarks?$/.test(t) ||
    /date\s+of\s+entry\s+into\s+service/.test(t)
  );
};

const isFormXIVEmploymentCardHeaderLabelText = (text) => {
  const t = String(text || '')
    .replace(/\s+/g, ' ')
    .replace(/\.+$/g, '')
    .trim()
    .toLowerCase();
  if (!t) return false;
  return (
    /name\s+and\s+address\s+(?:of|if)\s+(?:the\s+)?contractor/.test(t) ||
    /name\s+and\s+address\s+of\s+(?:the\s+)?establishment|establishment\s+in\s*\/?\s*under\s+which/.test(
      t
    ) ||
    /nature\s+and\s+location\s+of\s+work|nature\s+of\s+work\s+and\s+location/.test(t) ||
    /name\s+and\s+address\s+of\s+(?:the\s+)?principal\s+employer|(?:^|:\s*)principal\s+employer/.test(
      t
    )
  );
};

/** Canonical Form XIV header labels — always shown on PDF even when values are blank. */
const FORM_XIV_PDF_HEADER_FIELD_LABELS = [
  'Name and address if contractor',
  'Name and address of Establishment in/under which contract is carried on',
  'Nature and location of work',
  'Name and address of Principal Employer',
];

const matchFormXIVPdfHeaderFieldIndex = (text) => {
  const t = String(text || '')
    .replace(/\s+/g, ' ')
    .replace(/\.+$/g, '')
    .replace(/\s*:?\s*$/, '')
    .trim()
    .toLowerCase();
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

const FORM_XIV_PDF_STACKED_VALUE_COL_INDEX = 4; // Excel column E (1-based 5)
/** Gujarat Form XIV 2×2 header: right-hand box starts at Excel column G (0-based 6). */
const FORM_XIV_PDF_GJ_RIGHT_HALF_COL = 6;

/**
 * Collapse sparse Employment Card Excel columns into a single label|value pair per row
 * so Download PDF does not paint empty table columns.
 * Establishment header fields + signature stay outside the workman table box.
 */
const normalizeFormXIVEmploymentCardPdfMatrix = (rows, colCount, tableStartRow = 0) => {
  const src = Array.isArray(rows) ? rows : [];
  if (!src.length) {
    return { rows: src, colCount, tableStartRow, metaLines: null, formXIVFooterLines: [] };
  }

  // Horizontal Form X_RJ workman table — leave as-is.
  for (let r = 0; r < Math.min(src.length, 30); r += 1) {
    if (isFormXIVEmploymentCardTabularHeaderRow(src[r])) {
      return { rows: src, colCount, tableStartRow, metaLines: null, formXIVFooterLines: [] };
    }
  }

  const cellAt = (row, c) => String(row?.[c] ?? '').trim();

  const isFormXIVTitleLine = (text) => {
    const t = String(text || '').replace(/\s+/g, ' ').trim();
    if (!t) return false;
    return (
      /^form\s*xiv\b/i.test(t) ||
      /^employment\s+card$/i.test(t) ||
      /(?:see|vide)\s+rule\s*76/i.test(t) ||
      (/central\s*&\s*gujarat|gujarat\s+rules/i.test(t) && /rule\s*76/i.test(t))
    );
  };

  const isFormXIVPdfIgnorableValueText = (t) =>
    !t ||
    isFormXIVEmploymentCardHeaderLabelText(t) ||
    isFormXIVEmploymentCardWorkmanLabelText(t) ||
    /^\d{1,2}[\.\)]?$/.test(t) ||
    /^\.+$/.test(t) ||
    /^[_\s.-]{3,}$/.test(t) ||
    /signature\s+of\s+(?:the\s+)?contractor/i.test(t);

  /** Collect every header label on a row (GJ puts contractor|establishment on one row). */
  const collectHeaderHits = (row) => {
    const hits = [];
    const cells = Array.isArray(row) ? row : [];
    for (let c = 0; c < cells.length; c += 1) {
      const t = cellAt(cells, c);
      if (!t) continue;
      const headerIndex = matchFormXIVPdfHeaderFieldIndex(t);
      if (headerIndex < 0) continue;
      hits.push({
        col: c,
        headerIndex,
        half: c >= FORM_XIV_PDF_GJ_RIGHT_HALF_COL ? 'right' : 'left',
      });
    }
    return hits;
  };

  /** Split a value row into left/right 2×2 box halves (never join both into one string). */
  const splitRowHalves = (row) => {
    const cells = Array.isArray(row) ? row : [];
    const leftParts = [];
    const rightParts = [];
    for (let c = 0; c < cells.length; c += 1) {
      const t = cellAt(cells, c);
      if (isFormXIVPdfIgnorableValueText(t)) continue;
      if (c >= FORM_XIV_PDF_GJ_RIGHT_HALF_COL) rightParts.push(t);
      else leftParts.push(t);
    }
    return {
      left: leftParts.join(' ').trim(),
      right: rightParts.join(' ').trim(),
    };
  };

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
      if (/^\d{1,2}[\.\)]?$/.test(t)) {
        ordinal = t.replace(/\)/, '.');
        if (!/\.$/.test(ordinal)) ordinal = `${ordinal}.`;
        continue;
      }
      const combined = t.match(/^(\d{1,2})[\.\)]\s+(.+)$/);
      if (combined && isFormXIVEmploymentCardWorkmanLabelText(combined[2])) {
        ordinal = `${combined[1]}.`;
        label = combined[2].trim();
        labelCol = c;
        kind = 'workman';
        continue;
      }
      if (!label && isFormXIVEmploymentCardWorkmanLabelText(t)) {
        label = t.replace(/\.+$/g, '').trim();
        labelCol = c;
        kind = 'workman';
        continue;
      }
      const hdrIdx = matchFormXIVPdfHeaderFieldIndex(t);
      if (!label && hdrIdx >= 0) {
        label = FORM_XIV_PDF_HEADER_FIELD_LABELS[hdrIdx];
        labelCol = c;
        kind = 'header';
        headerIndex = hdrIdx;
      }
    }

    // Prefer stacked value column E when present (MP/KA Employment Card template).
    const stacked = cellAt(cells, FORM_XIV_PDF_STACKED_VALUE_COL_INDEX);
    if (
      stacked &&
      stacked !== label &&
      !isFormXIVEmploymentCardWorkmanLabelText(stacked) &&
      !isFormXIVEmploymentCardHeaderLabelText(stacked) &&
      !/^\d{1,2}[\.\)]?$/.test(stacked)
    ) {
      value = stacked;
    }

    if (!value) {
      for (let i = nonEmpty.length - 1; i >= 0; i -= 1) {
        const { c, t } = nonEmpty[i];
        if (c === labelCol) continue;
        if (/^\d{1,2}[\.\)]?$/.test(t)) continue;
        if (t === label) continue;
        if (isFormXIVEmploymentCardWorkmanLabelText(t) || isFormXIVEmploymentCardHeaderLabelText(t)) {
          continue;
        }
        // Dotted underline placeholders are not values.
        if (/^\.+$/.test(t) || /^[_\s.-]{3,}$/.test(t)) continue;
        value = t;
        break;
      }
    }

    if (!label && nonEmpty.length >= 1) {
      const first = nonEmpty[0].t;
      const hdrIdx = matchFormXIVPdfHeaderFieldIndex(first);
      if (hdrIdx >= 0) {
        label = FORM_XIV_PDF_HEADER_FIELD_LABELS[hdrIdx];
        kind = 'header';
        headerIndex = hdrIdx;
        value =
          value ||
          nonEmpty
            .slice(1)
            .map((x) => x.t)
            .filter(
              (t) =>
                t !== first &&
                !isFormXIVEmploymentCardHeaderLabelText(t) &&
                !/^\.+$/.test(t)
            )
            .join(' ');
      }
    }

    if (!label && !value) {
      return {
        left: nonEmpty.map((x) => x.t).join(' '),
        value: '',
        kind: 'title',
        headerIndex: -1,
      };
    }

    let left = label || '';
    if (kind === 'workman' && ordinal && left && !/^\d{1,2}[\.\)]/.test(left)) {
      left = `${ordinal} ${left}`.replace(/\.\s+/, '. ');
    } else if (ordinal && !left) {
      left = ordinal;
    }

    return { left, value, kind: kind || 'other', headerIndex };
  };

  // Already-compacted 2-col card: still lift headers + signature out of the table box.
  const looksCompact =
    Number(colCount) === 2 &&
    src.some((row) => isFormXIVEmploymentCardWorkmanLabelText(String(row?.[0] || '')));

  const titleMeta = [];
  const headerValues = FORM_XIV_PDF_HEADER_FIELD_LABELS.map(() => '');
  const headerSeen = FORM_XIV_PDF_HEADER_FIELD_LABELS.map(() => false);
  const workmanRows = [];
  const footerLines = [];
  const consumed = new Set();
  const pendingOrphanValues = [];

  const takeAheadHalfValues = (fromRow) => {
    let left = '';
    let right = '';
    for (let ar = fromRow + 1; ar < Math.min(src.length, fromRow + 8); ar += 1) {
      if (consumed.has(ar)) continue;
      const aheadRow = src[ar] || [];
      const aheadFilled = aheadRow.map((c) => String(c || '').trim()).filter(Boolean);
      if (!aheadFilled.length) continue;
      if (collectHeaderHits(aheadRow).length) break;
      if (
        aheadFilled.some(
          (t) =>
            isFormXIVEmploymentCardWorkmanLabelText(t) ||
            /^\d{1,2}[\.\)]\s+/.test(t) ||
            isFormXIVTitleLine(t) ||
            /signature\s+of\s+(?:the\s+)?contractor/i.test(t)
        )
      ) {
        break;
      }
      const halves = splitRowHalves(aheadRow);
      if (!halves.left && !halves.right) continue;
      if (halves.left) left = left ? `${left} ${halves.left}` : halves.left;
      if (halves.right) right = right ? `${right} ${halves.right}` : halves.right;
      consumed.add(ar);
    }
    return { left, right };
  };

  const assignHeaderValue = (idx, value) => {
    if (idx < 0 || idx >= headerValues.length) return;
    headerSeen[idx] = true;
    const text = String(value || '').trim();
    if (text && !headerValues[idx]) headerValues[idx] = text;
  };

  for (let r = 0; r < src.length; r += 1) {
    if (consumed.has(r)) continue;
    const row = src[r] || [];
    const filled = row.map((c) => String(c || '').trim()).filter(Boolean);
    if (!filled.length) continue;

    const blob = filled.join(' ').toLowerCase();
    if (isSystemGeneratedDocumentNote(filled.join(' '))) {
      continue;
    }
    if (/signature\s+of\s+(?:the\s+)?contractor/.test(blob)) {
      footerLines.push('Signature of the Contractor');
      continue;
    }

    // Compact 2-col rows already have label|value.
    if (looksCompact) {
      const left = String(row[0] || '').trim();
      const right = String(row[1] || '').trim();
      const hdrIdx = matchFormXIVPdfHeaderFieldIndex(left);
      if (hdrIdx >= 0) {
        assignHeaderValue(hdrIdx, right.replace(/^:\s*/, ''));
        continue;
      }
      if (isFormXIVEmploymentCardWorkmanLabelText(left) || /^\d{1,2}[\.\)]/.test(left)) {
        workmanRows.push([left, right]);
        continue;
      }
      if (isFormXIVTitleLine(left)) {
        titleMeta.push(left);
        continue;
      }
      if (right && !isFormXIVTitleLine(right)) {
        pendingOrphanValues.push(right);
      } else if (left && !isFormXIVTitleLine(left)) {
        pendingOrphanValues.push(left);
      }
      continue;
    }

    const headerHits = collectHeaderHits(row);
    if (headerHits.length > 0) {
      const sameRowHalves = splitRowHalves(row);
      const aheadHalves = takeAheadHalfValues(r);
      headerHits.forEach((hit) => {
        const fromSame =
          hit.half === 'right' ? sameRowHalves.right : sameRowHalves.left;
        const fromAhead = hit.half === 'right' ? aheadHalves.right : aheadHalves.left;
        let val = fromSame || fromAhead;
        // Stacked MP/KA: single label per row, value often only in column E (left half).
        if (!val && headerHits.length === 1) {
          const pair = splitStackedRow(row);
          val =
            pair.value ||
            (hit.half === 'right'
              ? aheadHalves.right || aheadHalves.left
              : aheadHalves.left || aheadHalves.right);
        }
        if (!val && pendingOrphanValues.length) val = pendingOrphanValues.shift();
        assignHeaderValue(hit.headerIndex, val);
      });
      continue;
    }

    const pair = splitStackedRow(row);

    if (pair.kind === 'workman') {
      // Flush leftover orphans into remaining empty header slots before workman body.
      while (pendingOrphanValues.length) {
        const emptyIdx = headerValues.findIndex((v) => !String(v || '').trim());
        if (emptyIdx < 0) break;
        assignHeaderValue(emptyIdx, pendingOrphanValues.shift());
      }
      workmanRows.push([pair.left, pair.value]);
      continue;
    }

    const titleLine = pair.left || filled.join(' ');
    if (isFormXIVTitleLine(titleLine)) {
      titleMeta.push(titleLine);
      continue;
    }

    // Orphan address / company text (GJ value boxes sit on rows without labels).
    // Keep left/right halves separate so Establishment / Principal Employer are not dropped.
    if (workmanRows.length === 0) {
      const halves = splitRowHalves(row);
      if (halves.left && halves.right) {
        pendingOrphanValues.push(halves.left, halves.right);
      } else {
        const orphan = String(halves.left || halves.right || pair.value || pair.left || '').trim();
        if (
          orphan &&
          !isFormXIVEmploymentCardHeaderLabelText(orphan) &&
          !isFormXIVEmploymentCardWorkmanLabelText(orphan)
        ) {
          pendingOrphanValues.push(orphan);
        }
      }
      continue;
    }

    // Anything after workman that isn't signature — ignore for PDF table.
  }

  while (pendingOrphanValues.length) {
    const emptyIdx = headerValues.findIndex((v) => !String(v || '').trim());
    if (emptyIdx < 0) break;
    assignHeaderValue(emptyIdx, pendingOrphanValues.shift());
  }

  const headerMetaLines = FORM_XIV_PDF_HEADER_FIELD_LABELS.map((label, idx) => {
    // Keep designation on the Nature line — never let "\nJunior Engineer" become a title band.
    const value = String(headerValues[idx] || '')
      .replace(/\s*\r?\n+\s*/g, ' ')
      .replace(/\s+/g, ' ')
      .trim();
    return value ? `${label}: ${value}` : `${label}:`;
  });

  const metaOut = [...titleMeta, ...headerMetaLines]
    .map((line) => String(line || '').replace(/\s*\r?\n+\s*/g, ' ').replace(/\s+/g, ' ').trim())
    .filter(Boolean);
  const uniqueFooter = [];
  footerLines.forEach((line) => {
    const key = String(line || '').toLowerCase();
    if (!key || uniqueFooter.some((x) => x.toLowerCase() === key)) return;
    uniqueFooter.push(line);
  });
  if (!uniqueFooter.length) uniqueFooter.push('Signature of the Contractor');

  return {
    rows: workmanRows.length ? workmanRows : [['1. Name of the Workman', '']],
    colCount: 2,
    tableStartRow: 0,
    metaLines: metaOut.length ? metaOut : null,
    formXIVFooterLines: uniqueFooter,
  };
};

const applyFormXIVKarnatakaPdfNormalization = (target, normalized) => {
  if (!target || !normalized) return target;
  target.rows = normalized.rows;
  target.colCount = normalized.colCount;
  target.tableStartRow = normalized.tableStartRow;
  if (Array.isArray(normalized.metaLines) && normalized.metaLines.length) {
    target.metaLines = normalized.metaLines;
  }
  target.formXIVFooterLines = Array.isArray(normalized.formXIVFooterLines)
    ? normalized.formXIVFooterLines
    : [];
  target.formXIVKALayout = true;
  return target;
};

const looksLikeForm14RajasthanPdfContext = (metaLines, rows, sheetName = '') => {
  const blob = [...(metaLines || []), ...(rows || []).slice(0, 12).flat(), sheetName || '']
    .join(' ')
    .toLowerCase();
  if (/form\s*14\b/.test(blob) && /rajasthan|hours\s+of\s+work|rule\s*22/.test(blob)) return true;
  if (/name\s+of\s+persons?\s+employ/.test(blob) && /whether\s+young\s+person/.test(blob)) return true;
  return /record\s+of\s+(?:the\s+)?hours\s+of\s+work/.test(blob) && /rajasthan|form\s*14\b/.test(blob);
};

/** Rajasthan Form D — Attendance Register (Place of work + day IN/OUT grid). */
const looksLikeFormDRajasthanPdfContext = (metaLines, rows, sheetName = '') => {
  const blob = [...(metaLines || []), ...(rows || []).slice(0, 18).flat(), sheetName || '']
    .join(' ')
    .toLowerCase();
  if (/gujarat|\b_gj\b|form[\s._-]*d[\s._-]*gj|relay\s+or\s+set/.test(blob)) return false;
  const hasFormD = /\bform[\s._-]*d\b/.test(blob);
  const hasRj = /rajasthan|\b_rj\b|form[\s._-]*d[\s._-]*rj|form_d_rj/.test(blob);
  const hasPlaceOfWork = /place\s+of\s+work/.test(blob);
  const hasSummaryDays = /summary\s+no\.?\s*of\s+days/.test(blob);
  const hasAttendance = /attendance\s+register|format\s+of\s+attendance/.test(blob);
  const hasOuterNote =
    /mines\s+only|underground|e\s*form\s+maintenance|weekly\s+off/.test(blob) &&
    /not\s+present|not\s+necessary|relay/.test(blob);
  if (hasRj && (hasFormD || hasAttendance)) return true;
  if (hasRj && hasPlaceOfWork && hasSummaryDays) return true;
  if (hasFormD && hasPlaceOfWork && hasSummaryDays && hasOuterNote) return true;
  if (hasPlaceOfWork && hasSummaryDays && hasAttendance) return true;
  return false;
};

/** Rajasthan Form 11 — Register of Employment. */
const looksLikeForm11RajasthanPdfContext = (metaLines, rows, sheetName = '') => {
  const blob = [...(metaLines || []), ...(rows || []).slice(0, 14).flat(), sheetName || '']
    .join(' ')
    .toLowerCase();
  if (/form\s*14\b/.test(blob) && /hours\s+of\s+work/.test(blob)) return false;
  if (/register\s+of\s+employment/.test(blob) && (/form\s*11\b/.test(blob) || /rajasthan|rule\s*22/.test(blob))) {
    return true;
  }
  if (
    /name\s+of\s+persons?\s+employ/.test(blob) &&
    /employment\s+commence/.test(blob) &&
    /total\s+hours\s+worked/.test(blob)
  ) {
    return true;
  }
  return false;
};

/** Rajasthan Form XIX — Register of Overtime templates often keep an empty column A. */
const looksLikeFormXIXRajasthanOvertimePdfContext = (metaLines, rows, sheetName = '') => {
  const blob = [...(metaLines || []), ...(rows || []).slice(0, 16).flat(), sheetName || '']
    .join(' ')
    .toLowerCase();
  return (
    /form\s*xix\b|form[\s._-]*xix[\s._-]/.test(blob) &&
    /register\s+of\s+overtime/.test(blob) &&
    (/rajasthan|\brj\b|rule\s*77/.test(blob) || /serial\s+no/.test(blob))
  );
};

const looksLikeFormAPdfContext = (metaLines, rows, sheetName = '') => {
  const blob = [...(metaLines || []), ...(rows || []).slice(0, 16).flat(), sheetName || '']
    .join(' ')
    .toLowerCase();
  return /\bform\s*a\b|form[\s._-]*a[\s._-]/.test(blob) && /employee\s+register/.test(blob);
};

/**
 * Form 11 RJ: Total hours worked during the month must be numeric or blank.
 * Clear leaked times (5PM / 9AM) and NIL from that column in the PDF matrix.
 */
const scrubForm11RajasthanTotalHoursPdfColumn = (rows, tableStartRow = 0, colCount = 0) => {
  if (!Array.isArray(rows) || rows.length === 0) return rows;
  const start = Math.max(0, Number(tableStartRow) || 0);
  const cols = Math.max(1, Number(colCount) || (rows[start] || []).length || 1);
  let hoursCol = -1;
  for (let r = start; r < Math.min(rows.length, start + 6); r += 1) {
    const row = rows[r] || [];
    for (let c = 0; c < cols; c += 1) {
      const t = String(row[c] || '')
        .replace(/\s+/g, ' ')
        .trim()
        .toLowerCase();
      if (/total\s+hours\s+worked/.test(t)) {
        hoursCol = c;
        break;
      }
    }
    if (hoursCol >= 0) break;
  }
  if (hoursCol < 0) return rows;

  let headerEnd = start;
  for (let r = start; r < Math.min(rows.length, start + 8); r += 1) {
    const blob = (rows[r] || []).join(' ').toLowerCase();
    if (/total\s+hours\s+worked|employment\s+commence|rest\s+interval|young\s+person/.test(blob)) {
      headerEnd = r;
    }
    const filled = (rows[r] || []).filter((c) => String(c || '').trim());
    if (filled.filter((c) => /^\d{1,2}$/.test(String(c).trim())).length >= 8) {
      headerEnd = Math.max(headerEnd, r);
    }
  }

  return rows.map((row, r) => {
    if (!Array.isArray(row) || r <= headerEnd) return row;
    if (isSystemGeneratedDocumentNoteRow(row)) return row;
    const next = row.slice();
    const raw = String(next[hoursCol] ?? '').trim();
    if (!raw) return next;
    // Keep pure numeric hours only.
    if (/^\d+(\.\d+)?$/.test(raw.replace(/,/g, ''))) return next;
    next[hoursCol] = '';
    return next;
  });
};

const isPureNilPdfText = (raw) => /^nill?$/i.test(String(raw || '').trim());

/** Form 25 TN muster / compensatory holidays register. */
const looksLikeForm25TamilNaduPdfContext = (metaLines, rows, sheetName = '') => {
  const blob = [...(metaLines || []), ...(rows || []).slice(0, 12).flat(), sheetName || '']
    .join(' ')
    .toLowerCase();
  if (/compensatory\s+holidays|muster\s+roll\s+and\s+register\s+of\s+compensatory/.test(blob)) {
    return true;
  }
  if (/form\s*(?:no\.?\s*)?[-–]?\s*25\b/.test(blob) && /tamil\s*nadu|daily\s+hours\s+of\s+work/.test(blob)) {
    return true;
  }
  return /form\s*(?:no\.?\s*)?[-–]?\s*25\b/.test(blob) && /prescribed\s+under\s+rules?\s*77/.test(blob);
};

/** Last 0-based column that has any non-empty cell in the matrix. */
const findLastContentColumnIndex = (rows) => {
  let last = -1;
  (rows || []).forEach((row) => {
    if (!Array.isArray(row)) return;
    for (let c = 0; c < row.length; c += 1) {
      if (String(row[c] || '').trim()) last = Math.max(last, c);
    }
  });
  return last;
};

/** 0-based Remarks column from the header band (Form 25 ends here). */
const findRemarksColumnIndex = (rows, scanRows = 20) => {
  let remarksIdx = -1;
  const limit = Math.min((rows || []).length, scanRows);
  for (let r = 0; r < limit; r += 1) {
    const row = rows[r] || [];
    for (let c = 0; c < row.length; c += 1) {
      if (/^remarks?$/i.test(String(row[c] || '').trim())) {
        remarksIdx = Math.max(remarksIdx, c);
      }
    }
  }
  return remarksIdx;
};

/**
 * Drop trailing blank columns (e.g. empty cells after Remarks on Form 25).
 * Optionally clamp to Remarks when that is the last meaningful header.
 */
const trimTrailingBlankPdfColumns = (rows, colCount, { clampToRemarks = false } = {}) => {
  let last = findLastContentColumnIndex(rows);
  if (clampToRemarks) {
    const remarksIdx = findRemarksColumnIndex(rows);
    if (remarksIdx >= 0) last = Math.min(last < 0 ? remarksIdx : last, remarksIdx);
  }
  const nextCount = Math.max(1, Math.min(colCount, last + 1));
  if (nextCount >= colCount) return { rows, colCount };
  const trimmed = (rows || []).map((row) => {
    const src = Array.isArray(row) ? row : [];
    const line = [];
    for (let c = 0; c < nextCount; c += 1) line.push(String(src[c] ?? ''));
    return line;
  });
  return { rows: trimmed, colCount: nextCount };
};

/**
 * Form 14 RJ templates often keep an empty column A left of "Name of Persons Employed".
 * Drop leading blank columns so the PDF table starts at the name column.
 */
const trimForm14RajasthanLeadingBlankPdfColumns = (rows, colCount, tableStartRow = 0) => {
  if (!Array.isArray(rows) || rows.length === 0 || colCount <= 1) {
    return { rows, colCount };
  }
  const startRow = Math.max(0, Number(tableStartRow) || 0);
  let lead = 0;

  // Prefer the official name-column header as the left edge of the table.
  outer: for (let r = startRow; r < Math.min(rows.length, startRow + 8); r += 1) {
    const row = rows[r] || [];
    for (let c = 0; c < colCount; c += 1) {
      const t = String(row[c] || '')
        .replace(/\s+/g, ' ')
        .trim()
        .toLowerCase();
      if (/name\s+of\s+persons?\s+employ/.test(t)) {
        lead = c;
        break outer;
      }
    }
  }

  // Fallback: first column with any non-footnote table-band content.
  if (lead <= 0) {
    for (let c = 0; c < colCount; c += 1) {
      let hasContent = false;
      for (let r = startRow; r < rows.length; r += 1) {
        const row = rows[r];
        if (!row) continue;
        if (isSystemGeneratedDocumentNoteRow(row) || isForm14RajasthanFootnoteRow(row)) continue;
        if (String(row[c] || '').trim()) {
          hasContent = true;
          break;
        }
      }
      if (hasContent) {
        lead = c;
        break;
      }
    }
  }

  if (lead <= 0) return { rows, colCount };

  const nextCount = Math.max(1, colCount - lead);
  const trimmed = rows.map((row) => {
    const src = Array.isArray(row) ? row : [];
    const line = [];
    for (let c = 0; c < nextCount; c += 1) line.push(String(src[lead + c] ?? ''));
    return line;
  });
  return { rows: trimmed, colCount: nextCount };
};

/** Drop leading spacer columns so the Rajasthan Form XIX table starts at Serial No. */
const trimFormXIXRajasthanLeadingBlankPdfColumns = (rows, colCount, tableStartRow = 0) => {
  if (!Array.isArray(rows) || rows.length === 0 || colCount <= 1) {
    return { rows, colCount };
  }
  const startRow = Math.max(0, Number(tableStartRow) || 0);
  let lead = 0;
  outer: for (let r = startRow; r < Math.min(rows.length, startRow + 8); r += 1) {
    const row = rows[r] || [];
    for (let c = 0; c < colCount; c += 1) {
      const text = String(row[c] || '').replace(/\s+/g, ' ').trim();
      if (/^(?:serial\s+no\.?|s\.?\s*no\.?)$/i.test(text)) {
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
};

/**
 * Form XI RJ Service Certificate templates keep an empty column A left of Serial No.
 * Drop leading blank columns so the PDF table starts at Serial No.
 */
const trimFormXIRajasthanLeadingBlankPdfColumns = (rows, colCount, tableStartRow = 0) => {
  if (!Array.isArray(rows) || rows.length === 0 || colCount <= 1) {
    return { rows, colCount };
  }
  const startRow = Math.max(0, Number(tableStartRow) || 0);
  let lead = 0;
  outer: for (let r = startRow; r < Math.min(rows.length, startRow + 10); r += 1) {
    const row = rows[r] || [];
    for (let c = 0; c < colCount; c += 1) {
      const text = String(row[c] || '')
        .replace(/\s+/g, ' ')
        .trim();
      if (/^(?:serial\s+no\.?|s\.?\s*no\.?|sl\.?\s*no\.?)$/i.test(text)) {
        lead = c;
        break outer;
      }
    }
  }
  // Fallback: first column with any table-band content (headers or data).
  if (lead <= 0) {
    for (let c = 0; c < colCount; c += 1) {
      let hasContent = false;
      for (let r = startRow; r < rows.length; r += 1) {
        if (String(rows[r]?.[c] || '').trim()) {
          hasContent = true;
          break;
        }
      }
      if (hasContent) {
        lead = c;
        break;
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
};

/**
 * Form XI RJ PDF: Rate of wages (piece-work particulars) stays blank when missing
 * or when the cell only has spilled designation / non-rate text.
 */
const sanitizeFormXIRajasthanServiceCertificatePdfDataRows = (rows, colCount, tableStartRow = 0) => {
  if (!Array.isArray(rows) || rows.length === 0 || colCount < 3) return rows;
  const startRow = Math.max(0, Number(tableStartRow) || 0);
  let headerRow = -1;
  let rateCol = -1;
  let natureCol = -1;
  for (let r = startRow; r < Math.min(rows.length, startRow + 10); r += 1) {
    const row = rows[r] || [];
    for (let c = 0; c < colCount; c += 1) {
      const text = String(row[c] || '')
        .replace(/\s+/g, ' ')
        .trim()
        .toLowerCase();
      if (!text) continue;
      if (/serial\s*no|s\.?\s*no/.test(text)) headerRow = r;
      if (/rate/.test(text) && /wage/.test(text)) {
        rateCol = c;
        headerRow = r;
      }
      if (/nature/.test(text) && /work/.test(text)) natureCol = c;
    }
    if (headerRow >= 0 && rateCol >= 0) break;
  }
  if (headerRow < 0 || rateCol < 0) return rows;

  const isNumericRate = (raw) => {
    const s = String(raw ?? '')
      .replace(/[,₹]/g, ' ')
      .replace(/\s+/g, ' ')
      .trim();
    if (!s) return false;
    if (/[a-z]/i.test(s) && !/^(rs\.?|inr)\b/i.test(s)) {
      const stripped = s.replace(/^(rs\.?|inr)\s*/i, '').trim();
      if (/[a-z]/i.test(stripped)) return false;
    }
    const num = Number(String(s).replace(/[^\d.]/g, ''));
    return Number.isFinite(num) && num > 0;
  };

  return rows.map((row, r) => {
    if (r <= headerRow) return row;
    const next = Array.from({ length: colCount }, (_, c) => String((row || [])[c] ?? ''));
    const filled = next.filter((v) => v.trim());
    const ordinalOnly =
      filled.length >= 3 && filled.every((v) => /^\d{1,2}$/.test(v.trim()));
    if (ordinalOnly) return next;

    const rateVal = String(next[rateCol] || '').trim();
    if (!rateVal) {
      next[rateCol] = '';
      return next;
    }
    const natureVal =
      natureCol >= 0 ? String(next[natureCol] || '').trim().toLowerCase() : '';
    if (
      !isNumericRate(rateVal) ||
      (natureVal && rateVal.toLowerCase() === natureVal) ||
      /engineer|manager|operator|supervisor/i.test(rateVal)
    ) {
      next[rateCol] = '';
    }
    return next;
  });
};

/** Drop the empty template column before the Rajasthan Form X workman table. */
const trimFormXRajasthanEmploymentLeadingBlankPdfColumns = (rows, colCount, tableStartRow = 0) => {
  if (!Array.isArray(rows) || rows.length === 0 || colCount <= 1) {
    return { rows, colCount };
  }
  const startRow = Math.max(0, Number(tableStartRow) || 0);
  let lead = 0;
  let foundWorkmanHeader = false;
  outer: for (let r = startRow; r < Math.min(rows.length, startRow + 8); r += 1) {
    const row = rows[r] || [];
    for (let c = 0; c < colCount; c += 1) {
      const text = String(row[c] || '').replace(/\s+/g, ' ').trim();
      if (/^name\s+of\s+the\s+workman$/i.test(text)) {
        lead = c;
        foundWorkmanHeader = true;
        break outer;
      }
    }
  }
  if (!foundWorkmanHeader) return { rows, colCount };
  const nextCount = Math.max(1, colCount - lead);
  const trimmed = rows.map((row) => {
    const src = Array.isArray(row) ? row : [];
    return Array.from({ length: nextCount }, (_, index) => String(src[lead + index] ?? ''));
  });
  return { rows: trimmed, colCount: nextCount };
};

/** Form X_RJ wage rate must be a positive amount — never designation / period text. */
const isFormXRajasthanEmploymentNumericWageRate = (raw) => {
  const s = String(raw ?? '')
    .replace(/[,₹]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
  if (!s) return false;
  if (/[a-z]/i.test(s) && !/^(rs\.?|inr)\b/i.test(s)) {
    // Allow "Rs 41412" / "INR 41412"; reject "Junior Engineer", dates, etc.
    const stripped = s.replace(/^(rs\.?|inr)\s*/i, '').trim();
    if (/[a-z]/i.test(stripped)) return false;
  }
  const num = Number(String(s).replace(/[^\d.]/g, ''));
  return Number.isFinite(num) && num > 0;
};

/**
 * Form X_RJ PDF: keep Wage rate blank when missing or spilled from Nature of employment.
 * Also clear Remarks when it is only a duplicate of Period of employment.
 */
const sanitizeFormXRajasthanEmploymentPdfDataRows = (rows, colCount, tableStartRow = 0) => {
  if (!Array.isArray(rows) || rows.length === 0 || colCount < 3) return rows;
  const startRow = Math.max(0, Number(tableStartRow) || 0);
  let headerRow = -1;
  let wageCol = -1;
  let natureCol = -1;
  let periodCol = -1;
  let remarksCol = -1;
  for (let r = startRow; r < Math.min(rows.length, startRow + 8); r += 1) {
    const row = rows[r] || [];
    for (let c = 0; c < colCount; c += 1) {
      const text = String(row[c] || '')
        .replace(/\s+/g, ' ')
        .trim()
        .toLowerCase();
      if (!text) continue;
      if (/name\s+of\s+the\s+workman/.test(text)) headerRow = r;
      if (/wage/.test(text) && /rate|piece|place|particular/.test(text) && !/wage\s+period/.test(text)) {
        wageCol = c;
        headerRow = r;
      }
      if (/nature/.test(text) && /employ|designat/.test(text)) natureCol = c;
      if (/tenure|period\s+of\s+employ/.test(text)) periodCol = c;
      if (/^remarks?$/.test(text) || /\bremarks?\b/.test(text)) remarksCol = c;
    }
    if (headerRow >= 0 && wageCol >= 0) break;
  }
  if (headerRow < 0 || wageCol < 0) return rows;

  return rows.map((row, r) => {
    if (r <= headerRow) return row;
    const next = Array.from({ length: colCount }, (_, c) => String((row || [])[c] ?? ''));
    // Skip pure ordinal rows (1..8).
    const numericOnly = next.filter((v) => v.trim()).every((v) => /^\d{1,2}$/.test(v.trim()));
    if (numericOnly && next.filter((v) => v.trim()).length >= 3) return next;

    const wageVal = String(next[wageCol] || '').trim();
    const natureVal = natureCol >= 0 ? String(next[natureCol] || '').trim() : '';
    if (
      wageVal &&
      (!isFormXRajasthanEmploymentNumericWageRate(wageVal) ||
        (natureVal && wageVal.toLowerCase() === natureVal.toLowerCase()))
    ) {
      next[wageCol] = '';
    }

    if (remarksCol >= 0 && periodCol >= 0) {
      const remarksVal = String(next[remarksCol] || '').trim();
      const periodVal = String(next[periodCol] || '').trim();
      if (remarksVal && periodVal && remarksVal.toLowerCase() === periodVal.toLowerCase()) {
        next[remarksCol] = '';
      }
    }
    return next;
  });
};

/** True when a Form 25 body row looks like a real employee (not leftover template shell). */
const isForm25TamilNaduEmployeePdfRow = (row) => {
  if (!Array.isArray(row)) return false;
  const filled = row.map((c) => String(c || '').trim()).filter(Boolean);
  if (!filled.length) return false;
  if (isSystemGeneratedDocumentNote(filled.join(' '))) return false;
  if (filled.some((v) => /^(P|A|WO|H|L|WOP|OD|SL|CL|EL|NH|FH)$/i.test(v))) return true;
  // Worker name / ID style tokens (not times, bare numbers, or shift labels alone).
  return filled.some((v) => {
    if (/^\d{1,2}:\d{2}/.test(v)) return false;
    if (/^\d{1,4}$/.test(v)) return false;
    if (/^(general\s+shift|scheme of shifts|rest(\s+interval)?|nil|n\/?a)$/i.test(v)) return false;
    if (/daily\s+hours|time at which|worker\s+identit|serial\s+number/i.test(v)) return false;
    return /[a-z]{2,}/i.test(v);
  });
};

/**
 * Form 25 templates often leave dozens of sample body rows — trim PDF after the last real employee.
 * Keep header band + footnote / system-note rows.
 */
const trimForm25TamilNaduPdfTrailingEmployeeRows = (rows, tableStartRow = 0) => {
  if (!Array.isArray(rows) || rows.length === 0) return rows;
  let headerEnd = Math.max(0, tableStartRow);
  for (let r = tableStartRow; r < Math.min(rows.length, tableStartRow + 8); r += 1) {
    const blob = (rows[r] || []).join(' ').toLowerCase();
    if (
      isForm25TamilNaduColHeaderBlob(blob) ||
      (/^\d{1,2}$/.test(String(rows[r]?.[0] || '').trim()) &&
        (rows[r] || []).filter((c) => /^\d{1,2}$/.test(String(c || '').trim())).length >= 10)
    ) {
      headerEnd = r;
    }
  }

  let lastEmp = -1;
  for (let r = headerEnd + 1; r < rows.length; r += 1) {
    if (isForm25TamilNaduEmployeePdfRow(rows[r])) lastEmp = r;
  }
  if (lastEmp < 0) return rows;

  const kept = [];
  for (let r = 0; r < rows.length; r += 1) {
    if (r <= lastEmp) {
      kept.push(rows[r]);
      continue;
    }
    // Preserve footer notes that sit below the employee block.
    if (isSystemGeneratedDocumentNoteRow(rows[r]) || isUnpaidAccumulationsFootnoteRow(rows[r])) {
      kept.push(rows[r]);
    }
  }
  return kept;
};

/** Drop rows that belong to a leaked Excel pivot (Row Labels / Count of MALE…). */
const stripLeakedPivotRowsFromPdfMatrix = (rows, tableStartRow = 0) => {
  if (!Array.isArray(rows) || rows.length === 0) return rows;
  const out = [];
  let dropping = false;
  for (let r = 0; r < rows.length; r += 1) {
    const blob = (rows[r] || []).join(' ').toLowerCase();
    if (
      r > tableStartRow &&
      (/^row\s+labels\b/.test(blob.trim()) ||
        (/row\s+labels/.test(blob) && /count of\s+mal/.test(blob)))
    ) {
      dropping = true;
    }
    if (dropping) {
      // Keep system note if it somehow sits after the pivot block.
      if (isSystemGeneratedDocumentNoteRow(rows[r])) {
        out.push(rows[r]);
      }
      continue;
    }
    out.push(rows[r]);
  }
  return out;
};

/** Form W admin band above the leaf table (employer / gender boxes / month-year). */
const isFormWAdminBandBlob = (blob) => {
  const t = String(blob || '').toLowerCase();
  if (!t || isWageRegisterColHeaderBlob(t)) return false;
  return (
    /name and address of the employer|name of the manager|registration certificate|wage period from/.test(
      t
    ) ||
    (/\bmen\b/.test(t) && /\bwomen\b/.test(t)) ||
    (/male young person|female young person/.test(t) && !/s\.?\s*no/.test(t)) ||
    (/^month\s*:/.test(t.trim()) || /\bmonth\s*:/.test(t)) ||
    (/^year\s*:/.test(t.trim()) || /\byear\s*:/.test(t))
  );
};

/** Gender count / Month-Year value row sitting between Form W admin labels and S.No. */
const isFormWAdminValueRow = (filled) => {
  if (!Array.isArray(filled) || filled.length === 0 || filled.length > 8) return false;
  const allShort = filled.every((t) => {
    const s = String(t || '').trim();
    if (!s) return true;
    if (/^\d{1,4}$/.test(s)) return true;
    if (/^(month|year)\s*:?\s*\S*$/i.test(s)) return true;
    // Month name alone (May / April 2024)
    if (/^(january|february|march|april|may|june|july|august|september|october|november|december)(\s+\d{4})?$/i.test(s)) {
      return true;
    }
    if (/^\d{4}$/.test(s)) return true;
    return false;
  });
  return allShort && filled.some((t) => String(t || '').trim());
};

const looksLikeFormWPdfContext = (metaLines, rows, tableStart = 0) => {
  const blob = [...(metaLines || []), ...(rows || []).slice(0, Math.min(rows.length, 14)).flat()]
    .map((t) => String(t || ''))
    .join(' ')
    .toLowerCase();
  if (looksLikeFormXVIIITamilNaduPdfContext(metaLines, rows, '', '')) return false;
  if (/wages[\s._-]*cum[\s._-]*muster|form[\s._-]*xviii(?![a-z])/.test(blob)) return false;
  const isFormW =
    /(?:^|[^a-z0-9])form[\s._-]*w(?:[^a-z0-9]|$)/.test(blob) ||
    (/register of wages/.test(blob) && /sub-rule\s*\(?\s*1\s*\)?\s*of\s*rule\s*\(?\s*16/.test(blob));
  if (!isFormW) return false;
  // Prefer Tamil Nadu Form W; also accept plain Form W templates without state text.
  return (
    /tamil[\s._-]*nadu|tamilnadu|sub-rule\s*\(?\s*1\s*\)?\s*of\s*rule\s*\(?\s*16/.test(blob) ||
    isWageRegisterColHeaderBlob(
      (rows || [])
        .slice(Math.max(0, tableStart), Math.min(rows.length, (tableStart || 0) + 4))
        .flat()
        .join(' ')
    )
  );
};

/**
 * Build a dense AOA matrix from a worksheet — all data cells, merges collapsed.
 */
const sheetToDenseMatrix = (worksheet, sheetName = 'Sheet', fileName = '') => {
  const aoa = XLSX.utils.sheet_to_json(worksheet, {
    header: 1,
    defval: '',
    raw: false,
    blankrows: true
  });
  if (!Array.isArray(aoa) || aoa.length === 0) {
    return { name: sheetName, rows: [], colCount: 0, metaLines: [], tableStartRow: 0 };
  }

  let rows = aoa.map((row) => {
    const src = Array.isArray(row) ? row : [];
    return src.map((c) => cellToText(c));
  });

  collapseMergedCellDuplicates(rows, worksheet['!merges'] || []);
  rows = rows.map((row) => dedupeRepeatedRowText(row));

  // Prefer Excel !ref width so empty trailing cols (signature etc.) stay in the PDF band.
  let refColCount = 0;
  try {
    const ref = worksheet['!ref'];
    if (ref) {
      const decoded = XLSX.utils.decode_range(ref);
      refColCount = Math.max(0, (decoded?.e?.c ?? -1) + 1);
    }
  } catch (_) {
    refColCount = 0;
  }

  let maxCol = 0;
  let lastContentRow = -1;
  for (let r = 0; r < rows.length && r < MAX_PDF_DATA_ROWS; r += 1) {
    // Footer note lives in col 0 after merge — do not shrink the table to that column.
    if (isSystemGeneratedDocumentNoteRow(rows[r])) {
      lastContentRow = Math.max(lastContentRow, r);
      continue;
    }
    let hasValue = false;
    for (let c = 0; c < rows[r].length; c += 1) {
      if (rows[r][c]) {
        hasValue = true;
        maxCol = Math.max(maxCol, c + 1);
      }
    }
    if (hasValue) lastContentRow = r;
  }

  if (lastContentRow < 0) {
    return { name: sheetName, rows: [], colCount: 0, metaLines: [], tableStartRow: 0 };
  }

  const endRow = Math.min(rows.length - 1, lastContentRow + MAX_TRAILING_EMPTY_AFTER_CONTENT);
  const earlyBlob = rows
    .slice(0, Math.min(rows.length, 14))
    .flat()
    .concat(sheetName || '')
    .join(' ')
    .toLowerCase();
  const isForm25Tn =
    /compensatory\s+holidays|form\s*(?:no\.?\s*)?[-–]?\s*25\b/.test(earlyBlob) &&
    (/prescribed\s+under\s+rules?\s*77|daily\s+hours\s+of\s+work|muster\s+roll/.test(earlyBlob) ||
      /form\s*(?:no\.?\s*)?[-–]?\s*25\b/.test(earlyBlob));

  // Wide day-grid forms (e.g. Form XXVI): keep a few empty trailing cols from !ref
  // (signature / termination). Form 25 ends at Remarks — never pad blank columns after it.
  if (
    !isForm25Tn &&
    refColCount > maxCol &&
    maxCol >= 20 &&
    refColCount - maxCol <= 8
  ) {
    maxCol = refColCount;
  }
  maxCol = Math.min(Math.max(maxCol, 1), MAX_PDF_COLS);

  // Clamp to Remarks when present so empty template columns after it are dropped.
  if (isForm25Tn) {
    const remarksIdx = findRemarksColumnIndex(rows);
    if (remarksIdx >= 0) maxCol = Math.min(maxCol, remarksIdx + 1);
  }

  const isFormITnSubsistence =
    /subsistence\s+allowance/.test(earlyBlob) ||
    /employees\s+placed\s+under\s+suspension/.test(earlyBlob) ||
    (/date\s+of\s+suspension/.test(earlyBlob) && /nature\s+of\s+offence/.test(earlyBlob));
  if (isFormITnSubsistence) {
    let signatureIdx = -1;
    for (let r = 0; r < Math.min(rows.length, 8); r += 1) {
      const row = rows[r] || [];
      for (let c = 0; c < row.length; c += 1) {
        const t = String(cellToText(row[c]) || '').toLowerCase();
        if (/signature\s+of\s+employee|postal\s+acknowledgement/.test(t)) {
          signatureIdx = Math.max(signatureIdx, c);
        }
      }
    }
    if (signatureIdx >= 0) maxCol = Math.min(maxCol, signatureIdx + 1);
    else maxCol = Math.min(maxCol, 12);
  }

  // Form XXIII GJ OT register is always 12 columns (A–L).
  if (looksLikeFormXXIIIGJOtPdfContext([], rows, sheetName, fileName)) {
    maxCol = Math.min(maxCol, 12);
  }

  const padded = [];
  for (let r = 0; r <= endRow; r += 1) {
    const line = [];
    for (let c = 0; c < maxCol; c += 1) {
      line.push(cellToText(rows[r]?.[c]));
    }
    padded.push(line);
  }

  // Gujarat Form M identity card — rebuild before generic meta/table splitting.
  if (looksLikeFormMGJGujaratPdfContext([], padded, sheetName)) {
    const normalized = normalizeFormMGJGujaratPdfMatrix(padded, maxCol, 0, []);
    return {
      name: sheetName,
      rows: normalized.rows,
      colCount: normalized.colCount,
      metaLines: normalized.metaLines,
      tableStartRow: normalized.tableStartRow,
      formMGJLayout: true,
      formMGJTitles: normalized.formMGJTitles,
      formMGJFields: normalized.formMGJFields,
      formMGJSignatures: normalized.formMGJSignatures,
    };
  }

  // Gujarat Form N Leave Book — three stacked sections, not one generic grid.
  if (looksLikeFormNGJGujaratPdfContext([], padded, sheetName)) {
    const normalized = normalizeFormNGJGujaratPdfMatrix(padded, maxCol, 0, []);
    return {
      name: sheetName,
      rows: normalized.rows,
      colCount: normalized.colCount,
      metaLines: normalized.metaLines,
      tableStartRow: normalized.tableStartRow,
      formNGJLayout: true,
      formNGJModel: normalized.formNGJModel,
    };
  }

  // Karnataka Form F leave register — identity + PART I + PART II, not one generic grid.
  if (looksLikeFormFKarnatakaPdfContext([], padded, sheetName, fileName)) {
    const normalized = normalizeFormFKarnatakaPdfMatrix(padded, maxCol, 0, []);
    return {
      name: sheetName,
      rows: normalized.rows,
      colCount: normalized.colCount,
      metaLines: normalized.metaLines,
      tableStartRow: normalized.tableStartRow,
      formFKALayout: true,
      formFKAModel: normalized.formFKAModel,
      fileName,
    };
  }

  // Karnataka Form P notice of maximum leave — letter + small table, not a generic grid.
  if (looksLikeFormPKarnatakaPdfContext([], padded, sheetName, fileName)) {
    const normalized = normalizeFormPKarnatakaPdfMatrix(padded, maxCol, 0, [], fileName);
    return {
      name: sheetName,
      rows: normalized.rows,
      colCount: normalized.colCount,
      metaLines: normalized.metaLines,
      tableStartRow: normalized.tableStartRow,
      formPKALayout: true,
      formPKAModel: normalized.formPKAModel,
      fileName,
    };
  }

  // Karnataka Form Q appointment order — 2-column Excel label|value, not stacked meta bands.
  if (looksLikeFormQKarnatakaPdfContext([], padded, sheetName)) {
    const normalized = normalizeFormQKarnatakaPdfMatrix(padded, maxCol, 0, []);
    return {
      name: sheetName,
      rows: normalized.rows,
      colCount: normalized.colCount,
      metaLines: normalized.metaLines,
      tableStartRow: normalized.tableStartRow,
      formQKALayout: true,
    };
  }

  // Split form title/meta (full-width once) from the data table
  let tableStartRow = 0;
  let metaLines = [];
  for (let r = 0; r < Math.min(padded.length, 20); r += 1) {
    const filled = padded[r].filter((c) => c);
    const blob = filled.join(' ').toLowerCase();
    const isFormXXAP = looksLikeFormXXAPPdfContext(
      [],
      padded.slice(0, Math.min(padded.length, 20)),
      sheetName
    );
    const isFormXXIAP = looksLikeFormXXIAPPdfContext(
      [],
      padded.slice(0, Math.min(padded.length, 20)),
      sheetName
    );
    const isFormXVIIAP = looksLikeFormXVIIAPPdfContext(
      [],
      padded.slice(0, Math.min(padded.length, 20)),
      sheetName
    );
    const isFormXVIAP = looksLikeFormXVIAPPdfContext(
      [],
      padded.slice(0, Math.min(padded.length, 24)),
      sheetName
    );
    const isFormXIXAP = isFormXIXAPWageSlipPdfContext(
      [],
      padded.slice(0, Math.min(padded.length, 24)),
      sheetName
    );
    const isFormXXIIIAPSE = looksLikeFormXXIIIAPSEPdfContext(
      metaLines,
      padded.slice(0, Math.min(padded.length, 24)),
      sheetName
    );
    const earlyFormXXIIIGJ = looksLikeFormXXIIIGJOtPdfContext(
      metaLines,
      padded.slice(0, Math.min(padded.length, 28)),
      sheetName,
      fileName
    );
    const earlyFormBGJ =
      looksLikeFormBGJGujaratPdfContext(
        metaLines,
        padded.slice(0, Math.min(padded.length, 24)),
        sheetName
      ) ||
      (/form\s*b|form_b_gj|format\s+of\s+wage\s+register/i.test(
        [...metaLines, sheetName || '', ...(padded[r] || [])].join(' ')
      ) &&
        /gujarat|_gj|wage\s+register|minimum\s+wages/i.test(
          [sheetName || '', ...padded.slice(0, 12).flat()].join(' ')
        ));
    if (
      earlyFormXXIIIGJ &&
      (isFormXXIIIGJOtTitleRow(padded[r]) ||
        isFormXXIIIGJOtAdminRow(padded[r]) ||
        isFormXXIIIGJOtPreambleRow(padded[r]))
    ) {
      if (!isFormXXIIIGJOtPreambleRow(padded[r])) {
        const unique = [];
        padded[r].forEach((cell) => {
          const text = String(cell || '').trim();
          if (text && !unique.includes(text)) unique.push(text);
        });
        unique.forEach((text) => {
          expandStatutoryMetaSegments(text).forEach((segment) => metaLines.push(segment));
        });
      }
      tableStartRow = r + 1;
      continue;
    }
    if (isFormXXIIIAPSE && (isFormXXIIIAPSETitleRow(padded[r]) || isFormXXIIIAPSEAdminRow(padded[r]))) {
      const unique = [];
      padded[r].forEach((cell) => {
        const text = String(cell || '').trim();
        if (text && !unique.includes(text)) unique.push(text);
      });
      unique.forEach((text) => {
        expandStatutoryMetaSegments(text).forEach((segment) => metaLines.push(segment));
      });
      tableStartRow = r + 1;
      continue;
    }

    if (
      (isFormXXAP || isFormXXIAP || isFormXVIIAP) &&
      (isFormXXAPAdministrativeRow(padded[r]) || isFormXVIIAdministrativeRow(padded[r]))
    ) {
      const unique = [];
      padded[r].forEach((cell) => {
        const text = String(cell || '').trim();
        if (text && !unique.includes(text)) unique.push(text);
      });
      unique.forEach((text) => {
        expandStatutoryMetaSegments(text).forEach((segment) => metaLines.push(segment));
      });
      tableStartRow = r + 1;
      continue;
    }

    // Form B GJ: Rate of Minimum Wages skill/value rows are header meta — never table start.
    if (earlyFormBGJ && isFormBGJMinimumWagesBandRow(padded[r])) {
      const unique = [];
      filled.forEach((t) => {
        if (!unique.includes(t)) unique.push(t);
      });
      unique.forEach((t) => {
        expandStatutoryMetaSegments(t).forEach((seg) => metaLines.push(seg));
      });
      tableStartRow = r + 1;
      continue;
    }

    // Form B GJ establishment / owner / LIN / principal / wage period (under min-wages box).
    if (
      earlyFormBGJ &&
      !isFormBGJWageRegisterColHeaderBlob(blob) &&
      (/name\s+of\s+(?:establishment|owner)|labour\s+identification|principal\s+employer|wage\s+period\s+from/i.test(
        blob
      ) ||
        (filled.length <= 2 &&
          filled.some((t) => t.length > 40 && /village|substation|pin\s*code|tal\.?|taluk/i.test(t))))
    ) {
      const unique = [];
      filled.forEach((t) => {
        if (!unique.includes(t)) unique.push(t);
      });
      unique.forEach((t) => {
        expandStatutoryMetaSegments(t).forEach((seg) => metaLines.push(seg));
      });
      tableStartRow = r + 1;
      continue;
    }

    const isColHeader =
      (earlyFormXXIIIGJ && isFormXXIIIGJOtTableHeaderRow(padded[r]) && filled.length >= 3) ||
      (isFormXXIIIAPSE && isFormXXIIIAPSETableHeaderRow(padded[r]) && filled.length >= 3) ||
      (isFormXXAPTableHeaderRow(padded[r]) && filled.length >= 5) ||
      (isFormXVIAPTableHeaderRow(padded[r]) && filled.length >= 3) ||
      (isFormXVIAP && isFormXVIIDateNumberRow(padded[r])) ||
      (earlyFormXXIIIGJ
        ? false
        : earlyFormBGJ
          ? isFormBGJWageRegisterColHeaderBlob(blob) && filled.length >= 3
          : isWageRegisterColHeaderBlob(blob) && filled.length >= 3) ||
      (isFormCLwfColHeaderBlob(blob) && filled.length >= 2) ||
      (isForm25TamilNaduColHeaderBlob(blob) && filled.length >= 3) ||
      (isForm14RajasthanColHeaderBlob(blob) && filled.length >= 2) ||
      (isFormXIXWageSlipColHeaderBlob(blob) && filled.length >= 3) ||
      (isFormTSEKarnatakaTableHeaderRow(padded[r]) && filled.length >= 3);
    const isNumberRow =
      filled.filter((c) => /^\d{1,2}$/.test(c)).length >= Math.max(6, filled.length * 0.6);

    if (isColHeader || isNumberRow) {
      tableStartRow = r;
      break;
    }

    // AP Form XIX stores each wage label/value as a short row, so do not
    // promote those rows into metadata before the AP normalizer can pair them.
    if (isFormXIXAP && filled.some(isFormXIXAPWageLabel)) {
      tableStartRow = r;
      break;
    }

    if (filled.length === 0) continue;

    // Form 25 festival-holiday boxes are just 1–5 — never print as underlined meta rows.
    if (
      filled.length <= 8 &&
      filled.every((t) => /^\d{1,2}$/.test(String(t || '').trim()))
    ) {
      tableStartRow = r + 1;
      continue;
    }

    // Form T KA: Month / Year, Establishment, Employer sit above the leaf header — keep as meta.
    if (
      looksLikeFormTSEKarnatakaPdfContext(metaLines, padded, sheetName) &&
      isFormTSEKarnatakaAdminRowBlob(blob) &&
      filled.length <= 4
    ) {
      const unique = [];
      filled.forEach((t) => {
        if (!unique.includes(t)) unique.push(t);
      });
      unique.forEach((t) => {
        expandStatutoryMetaSegments(t).forEach((seg) => metaLines.push(seg));
      });
      tableStartRow = r + 1;
      continue;
    }

    // Form W: employer / Men-Women / Month-Year sit above the leaf header — keep as meta.
    const alreadyFormWMeta = looksLikeFormWPdfContext(metaLines, [], 0);
    if (
      isFormWAdminBandBlob(blob) ||
      (alreadyFormWMeta && isFormWAdminValueRow(filled))
    ) {
      const unique = [];
      filled.forEach((t) => {
        if (!unique.includes(t)) unique.push(t);
      });
      unique.forEach((t) => {
        expandStatutoryMetaSegments(t).forEach((seg) => metaLines.push(seg));
      });
      tableStartRow = r + 1;
      continue;
    }

    // One (or few) meta values — show once as a line.
    // Do NOT promote multi-column rows of long labels (Form 14 RJ headers are all >20 chars).
    if (
      filled.length <= 2 ||
      (filled.length <= 3 && filled.every((t) => isFormMetaText(t) || t.length > 20))
    ) {
      const unique = [];
      filled.forEach((t) => {
        if (!unique.includes(t)) unique.push(t);
      });
      unique.forEach((t) => {
        expandStatutoryMetaSegments(t).forEach((seg) => metaLines.push(seg));
      });
      tableStartRow = r + 1;
      continue;
    }

    // Many columns but duplicated meta text already collapsed to col0
    if (filled.length === 1 && isFormMetaText(filled[0])) {
      expandStatutoryMetaSegments(filled[0]).forEach((seg) => metaLines.push(seg));
      tableStartRow = r + 1;
      continue;
    }

    tableStartRow = r;
    break;
  }

  let finalRows = padded;
  let formXIXAPLayout = false;
  let formXIXAPFooterLines = [];
  let formXIXGJHeaderLines = [];
  let formXIXGJLayout = false;
  let formXIXGJBoxedLayout = false;
  let formXIXKarnatakaLayout = false;
  let formXIXKABottomSection = null;
  let formXIVFooterLines = [];
  let formXIVKALayout = false;
  let formTKALayout = false;
  let formTKAAttendanceBand = null;
  let formTKAFooter = null;
  let formQKALayout = false;
  let formMGJLayout = false;
  let formMGJTitles = [];
  let formMGJFields = [];
  let formMGJSignatures = [];
  let formNGJLayout = false;
  let formNGJModel = null;
  let formFKALayout = false;
  let formFKAModel = null;
  let formPKALayout = false;
  let formPKAModel = null;
  if (isForm25Tn) {
    finalRows = trimForm25TamilNaduPdfTrailingEmployeeRows(padded, tableStartRow);
  }

  // Strip any pivot "Row Labels" block that leaked into a statutory sheet matrix.
  finalRows = stripLeakedPivotRowsFromPdfMatrix(finalRows, tableStartRow);

  let finalColCount = maxCol;
  if (looksLikeForm14RajasthanPdfContext(metaLines, finalRows, sheetName)) {
    const trimmedLead = trimForm14RajasthanLeadingBlankPdfColumns(
      finalRows,
      finalColCount,
      tableStartRow
    );
    finalRows = trimmedLead.rows;
    finalColCount = trimmedLead.colCount;
  }
  if (looksLikeFormXIXRajasthanOvertimePdfContext(metaLines, finalRows, sheetName)) {
    const trimmedLead = trimFormXIXRajasthanLeadingBlankPdfColumns(
      finalRows,
      finalColCount,
      tableStartRow
    );
    finalRows = trimmedLead.rows;
    finalColCount = trimmedLead.colCount;
  }
  if (looksLikeFormXXIIIGJOtPdfContext(metaLines, finalRows, sheetName, fileName)) {
    const normalizedGj = normalizeFormXXIIIGJOtPdfMatrix(
      finalRows,
      finalColCount,
      tableStartRow,
      metaLines
    );
    finalRows = normalizedGj.rows;
    finalColCount = normalizedGj.colCount;
    tableStartRow = normalizedGj.tableStartRow;
    if (Array.isArray(normalizedGj.metaLines)) metaLines = normalizedGj.metaLines;
  }
  if (looksLikeFormXIRajasthanServiceCertificatePdfContext(metaLines, finalRows, sheetName)) {
    const trimmedLead = trimFormXIRajasthanLeadingBlankPdfColumns(
      finalRows,
      finalColCount,
      tableStartRow
    );
    finalRows = sanitizeFormXIRajasthanServiceCertificatePdfDataRows(
      trimmedLead.rows,
      trimmedLead.colCount,
      tableStartRow
    );
    finalColCount = trimmedLead.colCount;
  }
  if (looksLikeFormXRajasthanEmploymentCardPdfContext(metaLines, finalRows, sheetName)) {
    const trimmedLead = trimFormXRajasthanEmploymentLeadingBlankPdfColumns(
      finalRows,
      finalColCount,
      tableStartRow
    );
    finalRows = sanitizeFormXRajasthanEmploymentPdfDataRows(
      trimmedLead.rows,
      trimmedLead.colCount,
      tableStartRow
    );
    finalColCount = trimmedLead.colCount;
  }
  if (looksLikeFormKGJGujaratPdfContext(metaLines, finalRows, sheetName)) {
    const trimmedLead = trimFormKGJGujaratLeadingBlankPdfColumns(
      finalRows,
      finalColCount,
      tableStartRow
    );
    finalRows = sanitizeFormKGJGujaratPdfDataRows(
      trimmedLead.rows,
      trimmedLead.colCount,
      tableStartRow
    );
    finalColCount = trimmedLead.colCount;
  }
  if (looksLikeFormLGJGujaratPdfContext(metaLines, finalRows, sheetName)) {
    const trimmedLead = trimFormLGJGujaratLeadingBlankPdfColumns(
      finalRows,
      finalColCount,
      tableStartRow
    );
    finalRows = sanitizeFormLGJGujaratPdfHeaderRows(
      trimmedLead.rows,
      trimmedLead.colCount,
      tableStartRow
    );
    finalColCount = trimmedLead.colCount;
  }
  if (looksLikeFormPGJGujaratPdfContext(metaLines, finalRows, sheetName)) {
    const trimmedLead = normalizeFormPGJGujaratPdfMatrix(
      finalRows,
      finalColCount,
      tableStartRow,
      metaLines
    );
    finalRows = trimmedLead.rows;
    finalColCount = trimmedLead.colCount;
  }
  if (looksLikeFormQMaharashtraPdfContext(metaLines, finalRows, sheetName, fileName)) {
    const trimmedLead = normalizeFormQMaharashtraPdfMatrix(
      finalRows,
      finalColCount,
      tableStartRow,
      metaLines
    );
    finalRows = trimmedLead.rows;
    finalColCount = trimmedLead.colCount;
  }
  if (looksLikeFormBTamilNaduPdfContext(metaLines, finalRows, sheetName, fileName)) {
    const normalized = normalizeFormBTamilNaduPdfMatrix(
      finalRows,
      finalColCount,
      tableStartRow,
      metaLines
    );
    finalRows = normalized.rows;
    finalColCount = normalized.colCount;
    tableStartRow = normalized.tableStartRow;
    if (Array.isArray(normalized.metaLines)) metaLines = normalized.metaLines;
  }
  if (looksLikeFormOGJGujaratPdfContext(metaLines, finalRows, sheetName)) {
    const trimmedLead = trimFormOGJGujaratLeadingBlankPdfColumns(
      finalRows,
      finalColCount,
      tableStartRow
    );
    finalRows = trimmedLead.rows;
    finalColCount = trimmedLead.colCount;
  }
  if (looksLikeFormXXIIIAPSEPdfContext(metaLines, finalRows, sheetName)) {
    let headerEnd = tableStartRow;
    for (let r = tableStartRow; r < Math.min(finalRows.length, tableStartRow + 6); r += 1) {
      if (
        isFormXXIIIAPSETableHeaderRow(finalRows[r]) ||
        isLikelyHeaderBandRow(finalRows[r], r)
      ) {
        headerEnd = r;
      }
    }
    const deductionsBand = findFormXXIIIAPSEDeductionsBand(
      finalRows,
      tableStartRow,
      finalColCount
    );
    if (deductionsBand) {
      headerEnd = Math.max(headerEnd, deductionsBand.childRow, deductionsBand.parentRow);
    }
    const trimmedXxiii = trimFormXXIIIAPSEEmptyPdfColumns(
      finalRows,
      tableStartRow,
      headerEnd,
      finalColCount
    );
    finalRows = trimmedXxiii.rows;
    finalColCount = trimmedXxiii.colCount;
  }
  if (looksLikeForm11RajasthanPdfContext(metaLines, finalRows, sheetName)) {
    finalRows = scrubForm11RajasthanTotalHoursPdfColumn(
      finalRows,
      tableStartRow,
      finalColCount
    );
  }
  if (
    looksLikeFormXIXKarnatakaPdfContext(metaLines, finalRows, sheetName) ||
    looksLikeFormXIXGJWageSlipPdfContext(metaLines, finalRows, sheetName) ||
    looksLikeFormXIXWageSlipPdfContext(metaLines, finalRows, sheetName)
  ) {
    const normalized = normalizeFormXIXWageSlipPdfByState(
      finalRows,
      finalColCount,
      tableStartRow,
      metaLines,
      sheetName
    );
    finalRows = normalized.rows;
    finalColCount = normalized.colCount;
    tableStartRow = normalized.tableStartRow;
    formXIXAPLayout = normalized.formXIXAPLayout === true;
    formXIXAPFooterLines = normalized.formXIXAPFooterLines || [];
    formXIXGJHeaderLines = Array.isArray(normalized.formXIXGJHeaderLines)
      ? normalized.formXIXGJHeaderLines
      : [];
    formXIXGJLayout = normalized.formXIXGJLayout === true;
    formXIXGJBoxedLayout = normalized.formXIXGJBoxedLayout === true;
    formXIXKarnatakaLayout = normalized.formXIXKarnatakaLayout === true;
    formXIXKABottomSection = normalized.formXIXKABottomSection || null;
    if (Array.isArray(normalized.metaLines)) metaLines = normalized.metaLines;
  }
  if (looksLikeFormXIVEmploymentCardPdfContext(metaLines, finalRows, sheetName)) {
    const normalized = looksLikeFormXIVKarnatakaPdfContext(metaLines, finalRows, sheetName)
      ? normalizeFormXIVKarnatakaPdfMatrix(finalRows, finalColCount, tableStartRow, metaLines)
      : normalizeFormXIVEmploymentCardPdfMatrix(finalRows, finalColCount, tableStartRow);
    finalRows = normalized.rows;
    finalColCount = normalized.colCount;
    tableStartRow = normalized.tableStartRow;
    if (Array.isArray(normalized.metaLines) && normalized.metaLines.length) {
      metaLines = normalized.metaLines;
    }
    if (Array.isArray(normalized.formXIVFooterLines)) {
      formXIVFooterLines = normalized.formXIVFooterLines;
    }
    if (normalized.formXIVKALayout === true) {
      formXIVKALayout = true;
    }
  }
  if (looksLikeFormBGJGujaratPdfContext(metaLines, finalRows, sheetName)) {
    const normalized = normalizeFormBGJGujaratPdfMatrix(
      finalRows,
      finalColCount,
      tableStartRow,
      metaLines
    );
    finalRows = normalized.rows;
    finalColCount = normalized.colCount;
    tableStartRow = normalized.tableStartRow;
    if (Array.isArray(normalized.metaLines)) metaLines = normalized.metaLines;
  }
  if (looksLikeFormTSEKarnatakaPdfContext(metaLines, finalRows, sheetName)) {
    const normalized = normalizeFormTSEKarnatakaPdfMatrix(
      finalRows,
      finalColCount,
      tableStartRow,
      metaLines
    );
    finalRows = normalized.rows;
    finalColCount = normalized.colCount;
    tableStartRow = normalized.tableStartRow;
    if (Array.isArray(normalized.metaLines)) metaLines = normalized.metaLines;
    formTKALayout = true;
    if (normalized.attendanceBand) formTKAAttendanceBand = normalized.attendanceBand;
    if (normalized.formTKAFooter) formTKAFooter = normalized.formTKAFooter;
  }
  if (looksLikeFormQKarnatakaPdfContext(metaLines, finalRows, sheetName)) {
    const normalized = normalizeFormQKarnatakaPdfMatrix(
      finalRows,
      finalColCount,
      tableStartRow,
      metaLines
    );
    finalRows = normalized.rows;
    finalColCount = normalized.colCount;
    tableStartRow = normalized.tableStartRow;
    if (Array.isArray(normalized.metaLines)) metaLines = normalized.metaLines;
    formQKALayout = true;
  }
  if (looksLikeFormMGJGujaratPdfContext(metaLines, finalRows, sheetName)) {
    const normalized = normalizeFormMGJGujaratPdfMatrix(
      finalRows,
      finalColCount,
      tableStartRow,
      metaLines
    );
    finalRows = normalized.rows;
    finalColCount = normalized.colCount;
    tableStartRow = normalized.tableStartRow;
    if (Array.isArray(normalized.metaLines)) metaLines = normalized.metaLines;
    formMGJLayout = true;
    formMGJTitles = normalized.formMGJTitles || [];
    formMGJFields = normalized.formMGJFields || [];
    formMGJSignatures = normalized.formMGJSignatures || [];
  }
  if (looksLikeFormNGJGujaratPdfContext(metaLines, finalRows, sheetName)) {
    const normalized = normalizeFormNGJGujaratPdfMatrix(
      finalRows,
      finalColCount,
      tableStartRow,
      metaLines
    );
    finalRows = normalized.rows;
    finalColCount = normalized.colCount;
    tableStartRow = normalized.tableStartRow;
    if (Array.isArray(normalized.metaLines)) metaLines = normalized.metaLines;
    formNGJLayout = true;
    formNGJModel = normalized.formNGJModel || null;
  }
  if (looksLikeFormFKarnatakaPdfContext(metaLines, finalRows, sheetName, fileName)) {
    const normalized = normalizeFormFKarnatakaPdfMatrix(
      finalRows,
      finalColCount,
      tableStartRow,
      metaLines
    );
    finalRows = normalized.rows;
    finalColCount = normalized.colCount;
    tableStartRow = normalized.tableStartRow;
    if (Array.isArray(normalized.metaLines)) metaLines = normalized.metaLines;
    formFKALayout = true;
    formFKAModel = normalized.formFKAModel || null;
  }
  if (looksLikeFormPKarnatakaPdfContext(metaLines, finalRows, sheetName, fileName)) {
    const normalized = normalizeFormPKarnatakaPdfMatrix(
      finalRows,
      finalColCount,
      tableStartRow,
      metaLines,
      fileName
    );
    finalRows = normalized.rows;
    finalColCount = normalized.colCount;
    tableStartRow = normalized.tableStartRow;
    if (Array.isArray(normalized.metaLines)) metaLines = normalized.metaLines;
    formPKALayout = true;
    formPKAModel = normalized.formPKAModel || null;
  }

  return {
    name: sheetName,
    rows: finalRows,
    colCount: finalColCount,
    metaLines,
    tableStartRow,
    formXIXAPLayout,
    formXIXAPFooterLines,
    formXIXGJHeaderLines,
    formXIXGJLayout,
    formXIXGJBoxedLayout,
    formXIXKarnatakaLayout,
    formXIXKABottomSection,
    formXIVFooterLines,
    formXIVKALayout,
    formTKALayout,
    attendanceBand: formTKAAttendanceBand,
    formTKAFooter,
    formQKALayout,
    formMGJLayout,
    formMGJTitles,
    formMGJFields,
    formMGJSignatures,
    formNGJLayout,
    formNGJModel,
    formFKALayout,
    formFKAModel,
    formPKALayout,
    formPKAModel,
    fileName,
  };
};

/** Fill any blank cells from ExcelJS (formulas / rich text) without re-duplicating merges. */
const enrichMatrixWithExcelJs = async (arrayBuffer, matrices) => {
  try {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(arrayBuffer);
    const byName = new Map(matrices.map((m) => [String(m.name || '').toLowerCase(), m]));

    workbook.eachSheet((ws) => {
      const key = String(ws.name || '').toLowerCase();
      // Never fall back to matrices[0] — that merged Sheet3 pivot rows into Form 15 / Form 25.
      const matrix = byName.get(key);
      if (!matrix || !matrix.rows.length) return;

      // Form XIV Employment Card is already compacted to label|value — never enrich by
      // Excel row index (that maps title/header rows onto workman rows and drops fields).
      const formXIVAlreadyNormalized =
        matrix.formXIVKALayout === true ||
        (looksLikeFormXIVEmploymentCardPdfContext(matrix.metaLines, matrix.rows, matrix.name) &&
          Number(matrix.colCount) === 2 &&
          matrix.rows.some(
            (row) =>
              matchFormXIVPdfHeaderFieldIndex(String(row?.[0] || '')) >= 0 ||
              isFormXIVEmploymentCardWorkmanLabelText(String(row?.[0] || ''))
          ));
      if (formXIVAlreadyNormalized) {
        return;
      }
      if (matrix.formMGJLayout === true) {
        return;
      }
      if (matrix.formNGJLayout === true) {
        return;
      }
      if (matrix.formXIXAPLayout === true || matrix.formXIXKarnatakaLayout === true) {
        return;
      }
      if (matrix.formTKALayout === true) {
        return;
      }
      if (matrix.formQKALayout === true) {
        return;
      }
      if (matrix.formFKALayout === true) {
        return;
      }
      if (matrix.formPKALayout === true) {
        return;
      }

      // Build merge non-master set
      const nonMaster = new Set();
      const merges = ws._merges || ws.model?.merges || {};
      const list = Array.isArray(merges) ? merges : Object.values(merges || {});
      list.forEach((m) => {
        const top = m?.top ?? m?.model?.top;
        const left = m?.left ?? m?.model?.left;
        const bottom = m?.bottom ?? m?.model?.bottom;
        const right = m?.right ?? m?.model?.right;
        if (top == null || left == null) return;
        for (let r = top; r <= bottom; r += 1) {
          for (let c = left; c <= right; c += 1) {
            if (r === top && c === left) continue;
            nonMaster.add(`${r},${c}`);
          }
        }
      });

      ws.eachRow({ includeEmpty: false }, (row, rowNumber) => {
        const r = rowNumber - 1;
        if (r < 0 || r >= MAX_PDF_DATA_ROWS) return;
        // Do not extend the matrix with rows from a mismatched / larger sheet.
        if (r >= matrix.rows.length) return;

        row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
          if (nonMaster.has(`${rowNumber},${colNumber}`)) return;
          const c = colNumber - 1;
          if (c < 0 || c >= matrix.colCount || c >= MAX_PDF_COLS) return;
          const text = cellToText(cell.value);
          if (!text) return;
          if (!matrix.rows[r][c]) matrix.rows[r][c] = text;
        });
      });

      // Form 25: ExcelJS enrich can widen past Remarks — trim blank trailing cols again.
      if (
        looksLikeForm25TamilNaduPdfContext(matrix.metaLines, matrix.rows, matrix.name)
      ) {
        const trimmed = trimTrailingBlankPdfColumns(matrix.rows, matrix.colCount, {
          clampToRemarks: true
        });
        matrix.rows = trimForm25TamilNaduPdfTrailingEmployeeRows(
          trimmed.rows,
          matrix.tableStartRow || 0
        );
        matrix.colCount = trimmed.colCount;
      }

      // Form 14 RJ: drop empty template column A left of the name column.
      if (looksLikeForm14RajasthanPdfContext(matrix.metaLines, matrix.rows, matrix.name)) {
        const trimmedLead = trimForm14RajasthanLeadingBlankPdfColumns(
          matrix.rows,
          matrix.colCount,
          matrix.tableStartRow || 0
        );
        matrix.rows = trimmedLead.rows;
        matrix.colCount = trimmedLead.colCount;
      }

      // Form XIX RJ: drop the empty template column before Serial No.
      if (looksLikeFormXIXRajasthanOvertimePdfContext(matrix.metaLines, matrix.rows, matrix.name)) {
        const trimmedLead = trimFormXIXRajasthanLeadingBlankPdfColumns(
          matrix.rows,
          matrix.colCount,
          matrix.tableStartRow || 0
        );
        matrix.rows = trimmedLead.rows;
        matrix.colCount = trimmedLead.colCount;
      }

      // Form XXIII GJ: rebuild August model (meta titles/admin + 12-col OT band).
      if (
        looksLikeFormXXIIIGJOtPdfContext(
          matrix.metaLines,
          matrix.rows,
          matrix.name,
          matrix.fileName || ''
        )
      ) {
        const normalizedGj = normalizeFormXXIIIGJOtPdfMatrix(
          matrix.rows,
          matrix.colCount,
          matrix.tableStartRow || 0,
          matrix.metaLines
        );
        matrix.rows = normalizedGj.rows;
        matrix.colCount = normalizedGj.colCount;
        matrix.tableStartRow = normalizedGj.tableStartRow;
        if (Array.isArray(normalizedGj.metaLines)) matrix.metaLines = normalizedGj.metaLines;
      }

      // Form XI RJ Service Certificate: drop the empty template column before Serial No.
      if (
        looksLikeFormXIRajasthanServiceCertificatePdfContext(
          matrix.metaLines,
          matrix.rows,
          matrix.name
        )
      ) {
        const trimmedLead = trimFormXIRajasthanLeadingBlankPdfColumns(
          matrix.rows,
          matrix.colCount,
          matrix.tableStartRow || 0
        );
        matrix.rows = sanitizeFormXIRajasthanServiceCertificatePdfDataRows(
          trimmedLead.rows,
          trimmedLead.colCount,
          matrix.tableStartRow || 0
        );
        matrix.colCount = trimmedLead.colCount;
      }

      // Form X RJ Employment Card: remove the empty first table column.
      if (looksLikeFormXRajasthanEmploymentCardPdfContext(matrix.metaLines, matrix.rows, matrix.name)) {
        const trimmedLead = trimFormXRajasthanEmploymentLeadingBlankPdfColumns(
          matrix.rows,
          matrix.colCount,
          matrix.tableStartRow || 0
        );
        matrix.rows = sanitizeFormXRajasthanEmploymentPdfDataRows(
          trimmedLead.rows,
          trimmedLead.colCount,
          matrix.tableStartRow || 0
        );
        matrix.colCount = trimmedLead.colCount;
      }

      // Form K GJ: drop the empty template column before Sr. No.
      if (
        looksLikeFormKGJGujaratPdfContext(
          matrix.metaLines,
          matrix.rows,
          matrix.name,
          matrix.fileName
        )
      ) {
        const trimmedLead = trimFormKGJGujaratLeadingBlankPdfColumns(
          matrix.rows,
          matrix.colCount,
          matrix.tableStartRow || 0
        );
        matrix.rows = sanitizeFormKGJGujaratPdfDataRows(
          trimmedLead.rows,
          trimmedLead.colCount,
          matrix.tableStartRow || 0
        );
        matrix.colCount = trimmedLead.colCount;
      }

      // Form L GJ: drop the empty template column before Sr. No.
      if (
        looksLikeFormLGJGujaratPdfContext(
          matrix.metaLines,
          matrix.rows,
          matrix.name,
          matrix.fileName
        )
      ) {
        const trimmedLead = trimFormLGJGujaratLeadingBlankPdfColumns(
          matrix.rows,
          matrix.colCount,
          matrix.tableStartRow || 0
        );
        matrix.rows = sanitizeFormLGJGujaratPdfHeaderRows(
          trimmedLead.rows,
          trimmedLead.colCount,
          matrix.tableStartRow || 0
        );
        matrix.colCount = trimmedLead.colCount;
      }

      // Form P GJ: drop the empty template column before Sr. No.
      if (
        looksLikeFormPGJGujaratPdfContext(
          matrix.metaLines,
          matrix.rows,
          matrix.name,
          matrix.fileName
        )
      ) {
        const trimmedLead = normalizeFormPGJGujaratPdfMatrix(
          matrix.rows,
          matrix.colCount,
          matrix.tableStartRow || 0,
          matrix.metaLines
        );
        matrix.rows = trimmedLead.rows;
        matrix.colCount = trimmedLead.colCount;
      }

      // Form Q MH: drop empty leading columns before Sr. No.
      if (
        looksLikeFormQMaharashtraPdfContext(
          matrix.metaLines,
          matrix.rows,
          matrix.name,
          matrix.fileName
        )
      ) {
        const trimmedLead = normalizeFormQMaharashtraPdfMatrix(
          matrix.rows,
          matrix.colCount,
          matrix.tableStartRow || 0,
          matrix.metaLines
        );
        matrix.rows = trimmedLead.rows;
        matrix.colCount = trimmedLead.colCount;
      }

      // Form B TN: keep Fine blank; Other deductions stay in the Other column.
      if (
        looksLikeFormBTamilNaduPdfContext(
          matrix.metaLines,
          matrix.rows,
          matrix.name,
          matrix.fileName
        )
      ) {
        const normalized = normalizeFormBTamilNaduPdfMatrix(
          matrix.rows,
          matrix.colCount,
          matrix.tableStartRow || 0,
          matrix.metaLines
        );
        matrix.rows = normalized.rows;
        matrix.colCount = normalized.colCount;
        matrix.tableStartRow = normalized.tableStartRow;
        if (Array.isArray(normalized.metaLines)) matrix.metaLines = normalized.metaLines;
      }

      // Form O GJ: drop the empty template column before Sr. No.
      if (
        looksLikeFormOGJGujaratPdfContext(
          matrix.metaLines,
          matrix.rows,
          matrix.name,
          matrix.fileName
        )
      ) {
        const trimmedLead = trimFormOGJGujaratLeadingBlankPdfColumns(
          matrix.rows,
          matrix.colCount,
          matrix.tableStartRow || 0
        );
        matrix.rows = trimmedLead.rows;
        matrix.colCount = trimmedLead.colCount;
      }

      // Form 11 RJ: blank non-numeric Total-hours cells (no leaked 5PM / 9AM).
      if (looksLikeForm11RajasthanPdfContext(matrix.metaLines, matrix.rows, matrix.name)) {
        matrix.rows = scrubForm11RajasthanTotalHoursPdfColumn(
          matrix.rows,
          matrix.tableStartRow || 0,
          matrix.colCount
        );
      }

      // Form XVI AP Muster Roll: trim trailing blank columns after Remarks.
      if (looksLikeFormXVIAPPdfContext(matrix.metaLines, matrix.rows, matrix.name)) {
        const trimmed = trimTrailingBlankPdfColumns(matrix.rows, matrix.colCount, {
          clampToRemarks: true
        });
        matrix.rows = trimmed.rows;
        matrix.colCount = trimmed.colCount;
      }

      // Form XIX Wage Slip: collapse sparse merges into the Excel 5-column band.
      // Gujarat XIX: headers above box, wage 1–7 inside, initials below.
      // Karnataka XIX: boxed label. value header + 5-col wage table (image-1 model).
      if (
        looksLikeFormXIXKarnatakaPdfContext(
          matrix.metaLines,
          matrix.rows,
          matrix.name,
          matrix.fileName
        ) ||
        looksLikeFormXIXGJWageSlipPdfContext(matrix.metaLines, matrix.rows, matrix.name) ||
        looksLikeFormXIXWageSlipPdfContext(matrix.metaLines, matrix.rows, matrix.name)
      ) {
        const normalized = normalizeFormXIXWageSlipPdfByState(
          matrix.rows,
          matrix.colCount,
          matrix.tableStartRow || 0,
          matrix.metaLines,
          matrix.name,
          matrix.fileName
        );
        if (normalized.formXIXKarnatakaLayout === true) {
          applyFormXIXKarnatakaPdfNormalization(matrix, normalized);
        } else {
          matrix.rows = normalized.rows;
          matrix.colCount = normalized.colCount;
          matrix.tableStartRow = normalized.tableStartRow;
          matrix.formXIXAPLayout = normalized.formXIXAPLayout === true;
          matrix.formXIXAPFooterLines = normalized.formXIXAPFooterLines || [];
          matrix.formXIXGJHeaderLines = Array.isArray(normalized.formXIXGJHeaderLines)
            ? normalized.formXIXGJHeaderLines
            : [];
          matrix.formXIXGJLayout = normalized.formXIXGJLayout === true;
          matrix.formXIXGJBoxedLayout = normalized.formXIXGJBoxedLayout === true;
          if (Array.isArray(normalized.metaLines)) matrix.metaLines = normalized.metaLines;
        }
      }

      // Form XIV Employment Card: single label|value columns (no empty table cells).
      if (looksLikeFormXIVEmploymentCardPdfContext(matrix.metaLines, matrix.rows, matrix.name)) {
        const normalized = looksLikeFormXIVKarnatakaPdfContext(
          matrix.metaLines,
          matrix.rows,
          matrix.name,
          matrix.fileName
        )
          ? normalizeFormXIVKarnatakaPdfMatrix(
              matrix.rows,
              matrix.colCount,
              matrix.tableStartRow || 0,
              matrix.metaLines
            )
          : normalizeFormXIVEmploymentCardPdfMatrix(
              matrix.rows,
              matrix.colCount,
              matrix.tableStartRow || 0
            );
        if (normalized.formXIVKALayout === true) {
          applyFormXIVKarnatakaPdfNormalization(matrix, normalized);
        } else {
          matrix.rows = normalized.rows;
          matrix.colCount = normalized.colCount;
          matrix.tableStartRow = normalized.tableStartRow;
          if (Array.isArray(normalized.metaLines) && normalized.metaLines.length) {
            matrix.metaLines = normalized.metaLines;
          }
          if (Array.isArray(normalized.formXIVFooterLines)) {
            matrix.formXIVFooterLines = normalized.formXIVFooterLines;
          }
        }
      }

      matrix.rows = stripLeakedPivotRowsFromPdfMatrix(matrix.rows, matrix.tableStartRow || 0);
    });
  } catch (err) {
    console.warn('ExcelJS enrich skipped:', err);
  }
  return matrices;
};

const collectWorkbookMatrices = async (arrayBuffer, label = 'Draft') => {
  const wb = XLSX.read(arrayBuffer, { type: 'array', cellDates: true, cellText: true });
  const matrices = [];
  (wb.SheetNames || []).forEach((name) => {
    const matrix = sheetToDenseMatrix(wb.Sheets[name], name || label, label);
    if (matrix.rows.length > 0 && matrix.colCount > 0) {
      matrices.push(matrix);
    }
  });
  if (!matrices.length) return [];
  return enrichMatrixWithExcelJs(arrayBuffer, matrices);
};

/**
 * Pivot / helper sheets (e.g. "Sheet3" with Row Labels + Count of MALE) must not become PDF pages.
 */
const looksLikeAuxiliaryOrPivotPdfSheet = (matrix) => {
  const name = String(matrix?.name || '').trim();
  const blob = [
    name,
    ...(matrix?.metaLines || []),
    ...(matrix?.rows || []).slice(0, 20).flat()
  ]
    .join(' ')
    .toLowerCase();
  if (/row\s+labels/.test(blob)) return true;
  if (/count of\s+mal/.test(blob) && /count of\s+fem/.test(blob)) return true;
  if (/sum of\s+gross\s+wages/.test(blob) && /sum of\s+basic/.test(blob)) return true;
  if (/pivot|count of male|count of female/.test(blob) && /sum of\s+(gross|basic|house\s+rent)/.test(blob)) {
    return true;
  }
  // Generic "Sheet3" style tab with no statutory form markers — only when clearly not a register.
  if (
    /^sheet\s*\d+$/i.test(name) &&
    !/form\s*(?:no\.?\s*)?[-–.]?\s*\d+|form\s*15|muster\s+roll|register of|prescribed under|compensatory|see\s+sub-?rule/i.test(
      blob
    )
  ) {
    if (/count of|sum of|row labels|gross wages|basic wage|admin assistant|designation/.test(blob)) {
      return true;
    }
  }
  return false;
};

/** Form 15 Part 1 TN — Register of Leave with Wages. */
const looksLikeForm15TamilNaduPdfContext = (metaLines, rows, sheetName = '') => {
  if (looksLikeFormFKarnatakaPdfContext(metaLines, rows, sheetName)) return false;
  if (looksLikeFormPKarnatakaPdfContext(metaLines, rows, sheetName)) return false;
  const blob = [...(metaLines || []), ...(rows || []).slice(0, 12).flat(), sheetName || '']
    .join(' ')
    .toLowerCase();
  if (/form\s*15|form-15|form_15/.test(blob) && (/leave|part\s*i\b|part\s*1\b/.test(blob) || /tamil/.test(blob))) {
    return true;
  }
  if (/register of leave/.test(blob) && (/earned leave|medical leave|see\s+sub-?rule/.test(blob))) {
    return true;
  }
  if (/earned leave/.test(blob) && /medical leave/.test(blob) && /name of the employee/.test(blob)) {
    return true;
  }
  return false;
};

/**
 * Catalog / download name says Form 15 Part 1 even when the Excel template still has FORM-X titles.
 * Used to rewrite PDF (and export) headings — never rewrite a true Form X download.
 */
const looksLikeForm15Part1PreferredContext = (preferredTitle = '', fileName = '', sheetName = '') => {
  const blob = [preferredTitle, fileName, sheetName]
    .filter(Boolean)
    .join(' ')
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, ' ');
  if (!blob) return false;
  if (/form\s*15/.test(blob) && /part\s*2|partii|part_2/.test(blob)) return false;
  if (/form\s*15/.test(blob) && /part\s*1|parti|part_1/.test(blob)) return true;
  if (
    /form\s*15/.test(blob) &&
    /register\s*of\s*leave\s*with\s*wages/.test(blob) &&
    !/part\s*2|partii|part_2/.test(blob)
  ) {
    return true;
  }
  return /form\s*15/.test(blob) && /tamil\s*nadu|tamilnadu/.test(blob) && /leave/.test(blob);
};

const FORM_15_PART1_PDF_TITLE = 'FORM-15';
const FORM_15_PART1_PDF_SUBTITLE = 'REGISTER OF LEAVE WITH WAGES';

/** Normalize dashes/spaces so "FORM–X" / "FORM X" still match. */
const normalizeFormTitleToken = (raw) =>
  String(raw || '')
    .replace(/[\u2010-\u2015\u2212\uFE58\uFE63\uFF0D]/g, '-')
    .replace(/\s+/g, ' ')
    .trim();

/** True for bare Form X only — never Form XV / XIX / XX / XXVI, etc. */
const isStandaloneFormXTitle = (raw) => {
  const t = normalizeFormTitleToken(raw);
  if (!t) return false;
  // FORM-X, FORM X, FORMX, Form - X
  if (/^form[\s._-]*x$/i.test(t)) return true;
  const compact = t.toLowerCase().replace(/[\s._-]+/g, '');
  return compact === 'formx';
};

/** Rewrite Form X template titles when the download is Form 15 Part 1. */
const rewriteForm15Part1PdfTitles = (titles) => {
  const list = Array.isArray(titles) ? [...titles] : [];
  const mapped = list
    .map((t) => {
      const raw = normalizeFormTitleToken(t);
      if (!raw) return '';
      // Drop Form X entirely — FORM-15 is ensured below (do not keep a mapped duplicate).
      if (isStandaloneFormXTitle(raw)) return '';
      if (/register\s+of\s+leave\s+and\s+social\s+security/i.test(raw)) {
        return FORM_15_PART1_PDF_SUBTITLE;
      }
      return raw;
    })
    .filter(Boolean)
    // Safety: strip any leftover Form X variant that survived mapping.
    .filter((t) => !isStandaloneFormXTitle(t));
  if (!mapped.some((t) => /form[\s._-]*15\b/i.test(t))) {
    mapped.unshift(FORM_15_PART1_PDF_TITLE);
  }
  if (!mapped.some((t) => /register\s+of\s+leave\s+with\s+wages/i.test(t))) {
    const formIdx = mapped.findIndex((t) => /form[\s._-]*15\b/i.test(t));
    mapped.splice(formIdx >= 0 ? formIdx + 1 : 0, 0, FORM_15_PART1_PDF_SUBTITLE);
  }
  return mapped.filter(
    (t, i, arr) => arr.findIndex((x) => normalizeMetaKey(x) === normalizeMetaKey(t)) === i
  );
};

/** Scrub Form X strings from meta lines before title/field extraction. */
const scrubForm15Part1MetaLines = (metaLines) =>
  (Array.isArray(metaLines) ? metaLines : [])
    .map((line) => {
      const raw = normalizeFormTitleToken(line);
      if (!raw) return '';
      if (isStandaloneFormXTitle(raw)) return '';
      if (/register\s+of\s+leave\s+and\s+social\s+security/i.test(raw)) {
        return FORM_15_PART1_PDF_SUBTITLE;
      }
      // Combined cell: "FORM-X REGISTER OF …" — drop the Form X token.
      return raw
        .replace(/\bform[\s._\-–—]*x\b(?!\s*[ivxlcdm0-9])/gi, '')
        .replace(/\s+/g, ' ')
        .trim();
    })
    .filter(Boolean);

/** Group banner labels: leave categories + Form 15 Part 2 wage/deduction bands. */
const isLeaveCategoryGroupLabel = (text) => {
  const t = String(text || '')
    .replace(/\s+/g, ' ')
    .trim();
  if (!t) return false;
  const n = t.toLowerCase();
  if (/^(earned|medical|other)\s+leave$/.test(n)) return true;
  if (/^maternity\s+benefits?$/.test(n)) return true;
  if (
    /\bleave\b/.test(n) &&
    /earned|medical|other/.test(n) &&
    !/beginning|availed|balance|period|during|earned\s+during|wages?/.test(n)
  ) {
    return true;
  }
  return false;
};

/** Form 15 Part 2 / Form W / Form XXVII style merged group banners (not leaf metric titles). */
const isWageDeductionGroupLabel = (text) => {
  const t = String(text || '')
    .replace(/\s+/g, ' ')
    .trim();
  if (!t) return false;
  const n = t.toLowerCase();
  if (/^deductions?$/.test(n)) return true;
  if (/earned\s+wages?\s+and\s+other\s+allowances/.test(n)) return true;
  if (/^earned\s+wages?$/.test(n)) return true;
  if (/^advances?$/.test(n)) return true;
  if (/^damages?\s*\/\s*fines?$/.test(n) || /^damages?\s+or\s+fines?$/.test(n)) return true;
  // Short "Leave Wages" group only — not the long leaf "Leave Wages (Earned…)".
  if (/^leave\s+wages?$/.test(n)) return true;
  if (/^other\s+allowances?$/.test(n)) return true;
  // Form XXVII TN Register of Wages group banners
  if (/^wages?\s+earned$/.test(n)) return true;
  if (isFormXVIIITamilNaduAmountOfWagesEarnedGroupLabel(t)) return true;
  if (/other\s+allowances?\s*\/\s*cash\s+payment|cash\s+payment\s+nature/.test(n)) return true;
  // Mid-tier "OTHER" over PT / Uniform Deposits (not "Other Deductions" leaf)
  if (/^other$/.test(n)) return true;
  if (isFormBTamilNaduAmountsDeductedGroupLabel(t)) return true;
  return false;
};

const isFormBRajasthanWageRateGroupLabel = (text) => {
  const normalized = String(text || '')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
  return /^rate\s+of\s+minimum\s+wages\s+and\s+since\s+the\s+date\b/.test(normalized);
};

const looksLikeFormBGJGujaratPdfContext = (metaLines, rows, sheetName = '', fileName = '') => {
  if (looksLikeFormPGJGujaratPdfContext(metaLines, rows, sheetName, fileName)) return false;
  const blob = [...(metaLines || []), ...(rows || []).slice(0, 22).flat(), sheetName || '', fileName || '']
    .join(' ')
    .toLowerCase();
  if (/rajasthan|\b_rj\b|form[\s._-]*b[\s._-]*rj|form_b_rj/.test(blob)) return false;
  if (/\bform\s*[-–.]?\s*p\b/.test(blob) && /muster[\s-]*roll\s+cum\s+wage/.test(blob)) return false;
  const hasFormB = /\bform\s*b\b|form[\s._-]*b[\s._-]|form_b_gj/.test(blob);
  const hasWageRegister =
    /format\s+of\s+wage\s+register/.test(blob) ||
    (/wage\s+register/.test(blob) && /rate\s+of\s+wage|net\s+payment/.test(blob));
  const hasGujarat = /gujarat|\b_gj\b|form[\s._-]*b[\s._-]*gj|form_b_gj/.test(blob);
  if (hasGujarat && (hasFormB || hasWageRegister)) return true;
  if (hasFormB && hasWageRegister) return true;
  return false;
};

const looksLikeFormBRajasthanPdfContext = (metaLines, rows, sheetName = '', fileName = '') => {
  const blob = [
    ...(metaLines || []),
    ...(rows || []).slice(0, 16).flat(),
    sheetName || '',
    fileName || '',
  ]
    .join(' ')
    .toLowerCase();
  if (/gujarat|\b_gj\b|form[\s._-]*b[\s._-]*gj|form_b_gj/.test(blob)) return false;
  if (/form[\s._-]*b[\s._-]*rj|form_b_rj|form\s*b\s*[-_]?\s*rajasthan/.test(blob)) return true;
  return (
    /\bform\s*b\b|form[\s._-]*b[\s._-]/.test(blob) &&
    /format\s+of\s+wage\s+register/.test(blob) &&
    (/rajasthan|minimum\s+wages/.test(blob) || /rate\s+of\s+minimum\s+wages/.test(blob))
  );
};

/** Excel-style Rate of Minimum Wages box (Form B GJ / RJ header, centered under titles). */
const extractFormBGJMinimumWagesBox = (metaLines = [], rows = [], tableStart = 0) => {
  const skillHeaders = ['Highly Skilled', 'Skilled', 'Semi-Skilled', 'Un-Skilled'];
  const rowLabels = ['Minimum Basic', 'DA', 'Overtime'];
  const box = {
    title: 'Rate of Minimum Wages',
    skills: skillHeaders,
    rows: rowLabels.map((label) => ({ label, values: ['', '', '', ''] })),
  };

  const sourceRows = [
    ...(Array.isArray(metaLines) ? metaLines.map((line) => [line]) : []),
    ...(Array.isArray(rows) ? rows.slice(0, Math.max(0, Number(tableStart) || 0) + 8) : []),
  ];

  let title = '';
  let skills = null;
  const dataRows = [];

  sourceRows.forEach((row) => {
    const cells = (Array.isArray(row) ? row : [row])
      .map((c) => String(c || '').replace(/\s+/g, ' ').trim())
      .filter(Boolean);
    if (!cells.length) return;
    const blob = cells.join(' ').toLowerCase();
    if (/rate\s+of\s+minimum\s+wages/.test(blob)) {
      title = cells.find((c) => /rate\s+of\s+minimum\s+wages/i.test(c)) || cells[0];
      return;
    }
    const skillHits = cells.filter((c) =>
      /^(highly\s*skilled|skilled|semi[-\s]?skilled|un[-\s]?skilled)$/i.test(c)
    );
    if (skillHits.length >= 3 && !skills) {
      skills = ['Highly Skilled', 'Skilled', 'Semi-Skilled', 'Un-Skilled'];
      return;
    }
    const labelIdx = cells.findIndex((c) =>
      /^(minimum\s+basic|da|dearness|overtime)$/i.test(c)
    );
    if (labelIdx >= 0) {
      const labelRaw = cells[labelIdx];
      let label = 'Minimum Basic';
      if (/^da\b|dearness/i.test(labelRaw)) label = 'DA';
      else if (/overtime/i.test(labelRaw)) label = 'Overtime';
      const values = cells.slice(labelIdx + 1, labelIdx + 5);
      while (values.length < 4) values.push('');
      dataRows.push({ label, values: values.slice(0, 4) });
    }
  });

  if (title) box.title = title;
  if (skills) box.skills = skills;
  if (dataRows.length) {
    box.rows = rowLabels.map((label) => {
      const found = dataRows.find((r) => r.label === label);
      return found || { label, values: ['', '', '', ''] };
    });
  }
  return box;
};

const isFormBGJMinimumWagesMetaText = (text) => {
  const t = String(text || '')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
  if (!t) return false;
  if (/rate\s+of\s+minimum\s+wages/.test(t)) return true;
  if (/^(highly\s*skilled|skilled|semi[-\s]?skilled|un[-\s]?skilled)$/.test(t)) return true;
  if (/^(minimum\s+basic|da|overtime)$/.test(t)) return true;
  return false;
};

/** True when a dense row is the Form B min-wages skill/value band (not the wage table). */
const isFormBGJMinimumWagesBandRow = (row) => {
  if (!Array.isArray(row)) return false;
  const cells = row.map((c) => String(c || '').replace(/\s+/g, ' ').trim()).filter(Boolean);
  if (!cells.length) return false;
  const blob = cells.join(' ').toLowerCase();
  if (/rate\s+of\s+minimum\s+wages/.test(blob)) return true;
  const skillHits = cells.filter((c) =>
    /^(highly\s*skilled|skilled|semi[-\s]?skilled|un[-\s]?skilled)$/i.test(c)
  ).length;
  if (skillHits >= 3) return true;
  if (
    cells.length <= 6 &&
    /^(minimum\s+basic|da|dearness|overtime)$/i.test(cells[0]) &&
    !/rate\s+of\s+wage|net\s+payment|sr\.?\s*no/i.test(blob)
  ) {
    return true;
  }
  return false;
};

/** Form B wage table header — not the Rate of Minimum Wages skill band. */
const isFormBGJWageRegisterColHeaderBlob = (blob) => {
  const t = String(blob || '').toLowerCase();
  if (!t) return false;
  if (/rate\s+of\s+minimum\s+wages|highly\s*skilled/.test(t) && !/sr\.?\s*no|rate\s+of\s+wage\b/.test(t)) {
    return false;
  }
  return (
    /sr\.?\s*no|s\.?\s*no/.test(t) &&
    (/employer|workman|worker|register|name|rate\s+of\s+wage|net\s+payment|days\s+worked/.test(t) ||
      /basic|hra|pf|esic/.test(t))
  );
};

/** Excel order for Form B Gujarat wage-register header fields (below min-wages box). */
const FORM_BGJ_PDF_FIELD_RANK = (text) => {
  const t = String(text || '')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
  if (!t) return 99;
  if (/name\s+of\s+establishment/.test(t) && !/principal/.test(t)) return 0;
  if (/name\s+of\s+owner/.test(t)) return 1;
  if (/labour\s+identification\s+no/.test(t) && /principal\s+employer/.test(t)) return 4;
  if (/labour\s+identification\s+no/.test(t)) return 2;
  if (/principal\s+employer/.test(t)) return 3;
  if (/wage\s+period/.test(t)) return 5;
  return 50;
};

const reorderFormBGJPdfHeaderFields = (fields) => {
  const list = Array.isArray(fields) ? fields.map((f) => String(f || '').trim()).filter(Boolean) : [];
  const seen = new Set();
  const unique = [];
  list.forEach((f) => {
    const key = f.toLowerCase();
    if (seen.has(key)) return;
    seen.add(key);
    unique.push(f);
  });
  return unique.sort((a, b) => FORM_BGJ_PDF_FIELD_RANK(a) - FORM_BGJ_PDF_FIELD_RANK(b));
};

/**
 * Pull Form B GJ Rate-of-Minimum-Wages + establishment bands out of the wage table
 * so they paint as header (box + unboxed fields below), matching Excel.
 */
const normalizeFormBGJGujaratPdfMatrix = (rows, colCount, tableStartRow = 0, metaLines = []) => {
  const src = Array.isArray(rows) ? rows.map((r) => (Array.isArray(r) ? [...r] : [])) : [];
  const meta = Array.isArray(metaLines) ? [...metaLines] : [];
  if (!src.length) {
    return { rows: src, colCount, tableStartRow, metaLines: meta };
  }

  const isEstablishmentHeaderBlob = (blob) =>
    /name\s+of\s+(?:establishment|owner)|labour\s+identification|principal\s+employer|wage\s+period\s+from/i.test(
      String(blob || '')
    );

  let start = Math.max(0, Number(tableStartRow) || 0);
  const promoted = [];
  let sawEstablishment = false;
  while (start < src.length) {
    const row = src[start] || [];
    const filled = row.map((c) => String(c || '').replace(/\s+/g, ' ').trim()).filter(Boolean);
    if (!filled.length) {
      start += 1;
      continue;
    }
    const blob = filled.join(' ');
    if (isFormBGJWageRegisterColHeaderBlob(blob) && filled.length >= 3) {
      break;
    }
    if (isFormBGJMinimumWagesBandRow(row)) {
      filled.forEach((t) => {
        expandStatutoryMetaSegments(t).forEach((seg) => {
          if (seg && !promoted.includes(seg) && !meta.includes(seg)) promoted.push(seg);
        });
      });
      start += 1;
      continue;
    }
    if (isEstablishmentHeaderBlob(blob)) {
      sawEstablishment = true;
      filled.forEach((t) => {
        expandStatutoryMetaSegments(t).forEach((seg) => {
          if (seg && !promoted.includes(seg) && !meta.includes(seg)) promoted.push(seg);
        });
      });
      start += 1;
      continue;
    }
    // Address continuation under establishment (no label) — only after a labeled est. row.
    if (
      sawEstablishment &&
      filled.length <= 2 &&
      filled[0].length > 24 &&
      /village|substation|pin\s*code|tal\.?|taluk/i.test(blob) &&
      !/sr\.?\s*no|rate\s+of\s+wage|net\s+payment/i.test(blob)
    ) {
      filled.forEach((t) => {
        expandStatutoryMetaSegments(t).forEach((seg) => {
          if (seg && !promoted.includes(seg) && !meta.includes(seg)) promoted.push(seg);
        });
      });
      start += 1;
      continue;
    }
    break;
  }

  return {
    rows: src,
    colCount,
    tableStartRow: start,
    metaLines: [...meta, ...promoted],
  };
};

/**
 * Form VI TN — merged banner above holiday date columns:
 * "Days, dates and months of the year on which National and Festival Holidays…"
 */
const isFormVIFestivalGroupLabel = (text) => {
  const t = String(text || '')
    .replace(/\s+/g, ' ')
    .trim();
  if (!t || t.length < 20) return false;
  const n = t.toLowerCase();
  if (/days,?\s+dates\s+and\s+months/.test(n)) return true;
  if (/enter\s+days\s+dates?/.test(n) && /months/.test(n)) return true;
  if (
    /national/.test(n) &&
    /festival/.test(n) &&
    /holiday/.test(n) &&
    (/section\s*3/.test(n) || /tamil\s+nadu\s+industrial\s+establishments/.test(n))
  ) {
    return true;
  }
  return false;
};

const isStatutoryGroupHeaderLabel = (text) =>
  isLeaveCategoryGroupLabel(text) ||
  isWageDeductionGroupLabel(text) ||
  isFormVIFestivalGroupLabel(text) ||
  isFormBRajasthanWageRateGroupLabel(text) ||
  isFormXXDateOfRecoveryGroupLabel(text) ||
  isFormXXIIIAPSEDeductionsGroupLabel(text) ||
  isFormBTamilNaduAmountsDeductedGroupLabel(text);

/** True when a leaf header marks the end of a Deductions / Leave Wages group span. */
const isGroupBandStopLeaf = (text) => {
  const n = String(text || '')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
  if (!n) return false;
  return (
    /^net\s+wages?$/.test(n) ||
    /^gross\s+wages?$/.test(n) ||
    /^date\s+of\s+payment$/.test(n) ||
    /^unpaid\s+accumulations?$/.test(n) ||
    /^remarks?$/.test(n) ||
    /^overtime\s+wages?$/.test(n) ||
    /^basic\s+wages?$/.test(n) ||
    /signature|thumb\s+impression|authori[sz]ed/i.test(n)
  );
};

/**
 * Leaf titles that belong outside a mid-level group (Advances / Damages),
 * so expansion does not swallow PF / Total Deductions / sibling bands.
 */
const isForeignLeafForGroup = (groupLabel, leafText) => {
  const g = String(groupLabel || '')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
  const n = String(leafText || '')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
  if (!n) return false;
  if (isGroupBandStopLeaf(n)) return true;
  if (/^advances?$/.test(g)) {
    return /damage|fine|any\s+other\s+deduction|total\s+deduction|provident|insurance|labour\s+welfare|leave\s+wage|gross\s+wage|overtime|basic\s+wage/.test(
      n
    );
  }
  if (/damage|fine/.test(g)) {
    return /advance|any\s+other\s+deduction|total\s+deduction|provident|insurance|labour\s+welfare|leave\s+wage|gross\s+wage|overtime|basic\s+wage/.test(
      n
    );
  }
  if (/^leave\s+wages?$/.test(g)) {
    return /gross\s+wage|overtime|basic\s+wage|deduction|provident|net\s+wage/.test(n);
  }
  if (/amount\s+of\s+wages?\s+earned|^wages?\s+earned$/.test(g)) {
    return /deduction|net\s+amount|net\s+wage|signature|thumb|initial|daily\s+rate|attendance|provident|unpaid|fines?/.test(
      n
    );
  }
  if (/other\s+allowances?\s*\/\s*cash\s+payment|cash\s+payment\s+nature/.test(g)) {
    return (
      /gross\s+wage|basic\s+wage|dearness|deduction|provident|net\s+wage|overtime/.test(n) ||
      isGroupBandStopLeaf(n)
    );
  }
  if (/^other$/.test(g)) {
    return (
      /fines?|other\s+deduction|total\s+deduction|provident|insurance|gross\s+wage|net\s+wage|signature|unpaid/.test(
        n
      ) || isGroupBandStopLeaf(n)
    );
  }
  if (/^deductions?$/.test(g)) {
    return isGroupBandStopLeaf(n);
  }
  if (/amounts?\s+deducted\s+during\s+the\s+month/.test(g)) {
    return /amount\s+actually\s+paid|balance\s+due|total\s+emolument|total\s+number/.test(n);
  }
  // Form VI festival-holiday banner — stop before Remarks / identity columns.
  if (
    /days,?\s+dates\s+and\s+months/.test(g) ||
    (/national/.test(g) && /festival/.test(g) && /holiday/.test(g) && g.length > 40)
  ) {
    return (
      /^remarks?$/.test(n) ||
      /^(?:s|sr|sl)\.?\s*no\.?$/.test(n) ||
      /employee\s+code|name\s+of\s+(?:the\s+)?(?:employee|worker)|d\.?o\.?j|date\s+of\s+join/.test(n)
    );
  }
  return false;
};

const headerBandHasLeafAt = (rows, fromRow, toRow, col) => {
  for (let r = fromRow; r <= toRow && r < rows.length; r += 1) {
    const t = String(rows[r]?.[col] || '').trim();
    if (!t) continue;
    if (isStatutoryGroupHeaderLabel(t)) continue;
    if (/^\(?\s*\d{1,2}\s*\)?$/.test(t)) continue; // column index row
    return true;
  }
  return false;
};

const readLeafTextAtCol = (rows, fromRow, toRow, col) => {
  for (let r = toRow; r >= fromRow && r < rows.length; r -= 1) {
    const t = String(rows[r]?.[col] || '')
      .replace(/\s+/g, ' ')
      .trim();
    if (!t) continue;
    if (isStatutoryGroupHeaderLabel(t)) continue;
    if (/^\(?\s*\d{1,2}\s*\)?$/.test(t)) continue;
    return t;
  }
  return '';
};

const expandStatutoryGroupBandEnd = (rows, groupRow, start, colCount, headerBandEnd, nextGroupStart, groupLabel) => {
  if (Number.isFinite(nextGroupStart) && nextGroupStart > start) {
    return nextGroupStart - 1;
  }
  let end = start;
  for (let c = start + 1; c < colCount; c += 1) {
    const group = String(rows[groupRow]?.[c] || '')
      .replace(/\s+/g, ' ')
      .trim();
    if (group) {
      if (isStatutoryGroupHeaderLabel(group)) break;
      // Non-group text on the same banner row ends the span.
      break;
    }
    const leaf = readLeafTextAtCol(rows, groupRow + 1, headerBandEnd, c);
    if (leaf && isForeignLeafForGroup(groupLabel, leaf)) break;
    if (leaf || headerBandHasLeafAt(rows, groupRow + 1, headerBandEnd, c) || c === start) {
      end = c;
      continue;
    }
    // Allow a single blank inside a merge; stop on a longer empty run.
    const nextLeaf = headerBandHasLeafAt(rows, groupRow + 1, headerBandEnd, c + 1);
    if (!nextLeaf) break;
    end = c;
  }
  return Math.max(start, end);
};

/**
 * Detect merged group header bands for leave registers and Form 15 Part 2
 * (Deductions / Advances / Damages / Leave Wages). Excel merges collapse to the
 * top-left cell — PDF must re-span those columns and center the label.
 */
const detectStatutoryGroupHeaderBands = (rows, tableStart, headerBandEnd, colCount) => {
  const bands = [];
  for (let r = tableStart; r <= headerBandEnd && r < rows.length; r += 1) {
    const found = [];
    for (let c = 0; c < colCount; c += 1) {
      const t = String(rows[r]?.[c] || '')
        .replace(/\s+/g, ' ')
        .trim();
      if (isStatutoryGroupHeaderLabel(t)) found.push({ start: c, label: t });
    }
    if (!found.length) continue;
    found.forEach((item, i) => {
      const nextStart = i + 1 < found.length ? found[i + 1].start : null;
      const end = expandStatutoryGroupBandEnd(
        rows,
        r,
        item.start,
        colCount,
        headerBandEnd,
        nextStart,
        item.label
      );
      if (end > item.start) {
        bands.push({
          labelRow: r,
          start: item.start,
          end,
          label: item.label
        });
      }
    });
  }
  return bands;
};

/** @deprecated alias — leave-only detection now covered by detectStatutoryGroupHeaderBands. */
const detectLeaveCategoryBands = (rows, tableStart, headerBandEnd, colCount) =>
  detectStatutoryGroupHeaderBands(rows, tableStart, headerBandEnd, colCount).filter((b) =>
    isLeaveCategoryGroupLabel(b.label)
  );

/**
 * When the workbook contains a primary statutory register, drop unrelated sheets
 * so PDF is only the form — not Sheet2/Sheet3 pivots that inflate page count.
 */
const filterStatutoryPdfMatrices = (matrices) => {
  const list = Array.isArray(matrices) ? matrices.filter(Boolean) : [];
  if (!list.length) return list;

  // Always drop pivot / helper sheets first.
  const withoutAux = list.filter((m) => !looksLikeAuxiliaryOrPivotPdfSheet(m));
  const pool = withoutAux.length > 0 ? withoutAux : list;

  const form25 = pool.filter((m) =>
    looksLikeForm25TamilNaduPdfContext(m.metaLines, m.rows, m.name)
  );
  if (form25.length > 0) return form25;

  const form15 = pool.filter((m) =>
    looksLikeForm15TamilNaduPdfContext(m.metaLines, m.rows, m.name)
  );
  if (form15.length > 0) return form15;

  // Prefer real form/register sheets over bare SheetN tabs.
  const formish = pool.filter((m) => {
    const blob = [...(m.metaLines || []), ...(m.rows || []).slice(0, 8).flat(), m.name || '']
      .join(' ')
      .toLowerCase();
    return (
      /\bform\b|register of|muster roll|prescribed under|see\s+sub-?rule|earned leave/i.test(blob) &&
      !looksLikeAuxiliaryOrPivotPdfSheet(m)
    );
  });
  if (formish.length > 0) return formish;

  return pool;
};

const rowHasContent = (row) =>
  Array.isArray(row) && row.some((c) => String(c || '').trim() !== '');

const isLikelyHeaderBandRow = (row, rowIndex) => {
  if (!rowHasContent(row)) return false;
  const filled = row.filter((c) => String(c || '').trim());
  const blob = filled.join(' ').toLowerCase();
  if (rowIndex < 20 && isWageRegisterColHeaderBlob(blob) && filled.length >= 3) {
    return true;
  }
  if (rowIndex < 20 && isForm14RajasthanColHeaderBlob(blob) && filled.length >= 2) {
    return true;
  }
  if (
    rowIndex < 20 &&
    /designation|wage|attendance|daily hours|amount of wages|net amount|deductions|gross wages|net wages/.test(
      blob
    ) &&
    filled.length >= 3
  ) {
    return true;
  }
  const nums = filled.filter((c) => /^\d{1,2}$/.test(String(c).trim()));
  return nums.length >= 8 && nums.length >= filled.length * 0.7;
};

const isPurePdfNumericText = (raw) => {
  const numericRaw = String(raw || '')
    .replace(/,/g, '')
    .replace(/^[₹$€£]\s?/, '')
    .replace(/^\((.+)\)$/, '-$1')
    .replace(/%$/, '')
    .trim();
  return /^-?\d+(\.\d+)?$/.test(numericRaw);
};

/** Form 26 / 26-A / Form 1 TN style "Nil of the month" / "Nil for the month of …" cell text. */
const isNilOfTheMonthPdfText = (text) => {
  const t = String(text || '').trim();
  return /^nill?\s+of\s+the\s+month\b/i.test(t) || /^nil\s+for\s+the\s+month\b/i.test(t);
};

/** Form 1 TN template spill: signature block / payment date in the amount column. */
const isNilOfTheMonthIgnorableSpillPdfText = (text) => {
  const raw = String(text || '')
    .replace(/_x000d_/gi, '\n')
    .replace(/\r\n/g, '\n')
    .replace(/\r/g, '\n')
    .trim();
  if (!raw) return true;
  const flat = raw.replace(/\s+/g, ' ').trim();
  if (/^for\s*\(/i.test(flat)) return true;
  if (/authorised\s+signatory|authorized\s+signatory/i.test(flat)) return true;
  if (/signature\s+of\s+employer/i.test(flat)) return true;
  if (/manager\s*\/\s*authori[sz]ed\s+person/i.test(flat)) return true;
  // Bare calendar date (31-07-2026) left from template / pay-run stamp — do not show.
  if (/^\d{1,2}[-/.]\d{1,2}[-/.]\d{2,4}$/.test(flat)) return true;
  // Multiline spill: leftover template date + For (… + Authorised Signatory.
  if (
    /\d{1,2}[-/.]\d{1,2}[-/.]\d{2,4}/.test(flat) &&
    (/for\s*\(/i.test(flat) || /authori[sz]ed\s+signatory/i.test(flat) || /signature\s+of\s+employer/i.test(flat))
  ) {
    return true;
  }
  return false;
};

/**
 * Extra values allowed on a nil-of-month data row (Form 1 TN puts serial + month
 * beside "Nill of the month" instead of a single merged cell).
 */
const isNilOfTheMonthCompanionPdfText = (text) => {
  const t = String(text || '').trim();
  if (!t) return true;
  if (isNilOfTheMonthIgnorableSpillPdfText(t)) return true;
  if (isNilOfTheMonthPdfText(t)) return true;
  if (/^nill?$/i.test(t)) return true;
  if (/^\d{1,4}$/.test(t)) return true; // Sl.No
  // Aug 2026 / August 2026
  if (/^(jan|feb|mar|apr|may|jun|jul|aug|sep|sept|oct|nov|dec)[a-z]*\.?\s+\d{4}$/i.test(t)) {
    return true;
  }
  if (/^\d{1,2}[-/]\d{4}$/.test(t) || /^\d{4}[-/]\d{1,2}$/.test(t)) return true;
  return false;
};

const NIL_OF_THE_MONTH_PDF_DISPLAY = 'Nil of the Month';

const NIL_OF_THE_MONTH_YEAR_RE =
  /\b((?:jan|feb|mar|apr|may|jun|jul|aug|sep|sept|oct|nov|dec)[a-z]*\.?)\s+(\d{4})\b/i;

/** Normalize "Aug 2026" / "August 2026" → "Aug 2026". */
const formatNilOfTheMonthYearLabel = (monthToken, yearToken) => {
  const abbr = String(monthToken || '')
    .replace(/\./g, '')
    .slice(0, 3);
  const y = String(yearToken || '').trim();
  if (!abbr || !/^\d{4}$/.test(y)) return '';
  return `${abbr.charAt(0).toUpperCase()}${abbr.slice(1).toLowerCase()} ${y}`;
};

/**
 * Prefer "Nil of the Month Aug 2026" when the period is on the nil line or a companion cell.
 */
const formatNilOfTheMonthPdfDisplay = (nilText, filledTexts = []) => {
  const fromNil = String(nilText || '').match(NIL_OF_THE_MONTH_YEAR_RE);
  if (fromNil) {
    const label = formatNilOfTheMonthYearLabel(fromNil[1], fromNil[2]);
    return label ? `${NIL_OF_THE_MONTH_PDF_DISPLAY} ${label}` : NIL_OF_THE_MONTH_PDF_DISPLAY;
  }
  for (const t of filledTexts) {
    if (isNilOfTheMonthPdfText(t)) continue;
    const m = String(t || '').trim().match(
      /^(jan|feb|mar|apr|may|jun|jul|aug|sep|sept|oct|nov|dec)[a-z]*\.?\s+(\d{4})$/i
    );
    if (!m) continue;
    const label = formatNilOfTheMonthYearLabel(m[1], m[2]);
    if (label) return `${NIL_OF_THE_MONTH_PDF_DISPLAY} ${label}`;
  }
  return NIL_OF_THE_MONTH_PDF_DISPLAY;
};

/**
 * Detect a data row that is a nil-month notice (after Excel merges collapse).
 * Returns a full-width span so PDF paints one merged, center-aligned box.
 * Form XXIX / Form 26: "1 | Nill of the month | Aug 2026 | …" → "Nil of the Month Aug 2026".
 * Form 1 TN subsistence keeps a month-end payment date in its own column, so that row
 * is not merged (the date must stay visible).
 */
const resolveNilOfTheMonthPdfSpan = (row, colCount, rowIndex, headerBandEnd) => {
  if (rowIndex <= headerBandEnd || colCount < 2) return null;
  const filled = [];
  let nilText = '';
  for (let c = 0; c < colCount; c += 1) {
    const t = String(row?.[c] ?? '').trim();
    if (!t) continue;
    if (isNilOfTheMonthIgnorableSpillPdfText(t)) continue;
    filled.push(t);
    if (isNilOfTheMonthPdfText(t)) nilText = t;
  }
  if (!filled.length || !nilText) return null;
  if (!filled.every((t) => isNilOfTheMonthCompanionPdfText(t))) return null;
  return {
    start: 0,
    end: colCount - 1,
    text: formatNilOfTheMonthPdfDisplay(nilText, filled)
  };
};

/** True when a column header is a serial / Sl. No. style label. */
const isPdfSerialNumberHeader = (headerText) => {
  const h = String(headerText || '')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
  if (!h) return false;
  if (/^\(?\s*\d{1,2}\s*\)?$/.test(h)) return false;
  return /^(?:s|sr|si|sl)\.?\s*no\.?$|^serial\s*(?:no\.?|number)\b|^sl\.?\s*no\b/i.test(h);
};

/**
 * Weight a column from its header name (+ data length) so PDF widths
 * follow column titles automatically for every statutory form.
 */
const statutoryHeaderColumnWeight = (headerText, maxDataLen = 0, colCount = 12) => {
  const h = String(headerText || '')
    .replace(/\s+/g, ' ')
    .trim();
  const lower = h.toLowerCase();
  const headerLen = h.length;
  const dataLen = Math.max(0, Number(maxDataLen) || 0);
  const wideForm = colCount > 16;

  if (isPdfSerialNumberHeader(h)) return Math.max(10, Math.min(12, 5 + Math.min(dataLen, 3)));

  // Short categorical headers — keep a floor so Sex/Age/Photo do not collapse.
  if (/^(sex|gender|age|photo|male|female)$/i.test(lower)) return 5.5;
  if (/^(nil|remarks?)$/i.test(lower)) return Math.max(5.5, Math.min(8, 4 + dataLen * 0.25));

  // ESI / insurance / id
  if (/insurance\s*no|esi\s*no|identification/i.test(lower)) {
    return Math.max(6.5, Math.min(9, 5 + Math.min(dataLen, 8) * 0.35));
  }

  // Date / time short labels
  if (/^(date|time)\b/i.test(lower) && headerLen <= 28) {
    return Math.max(5.5, Math.min(8.5, 4.5 + Math.min(dataLen, 12) * 0.25));
  }

  // Name / address identity columns
  if (/name.*address|address.*name|name of|injured person|name\s*&\s*address|name of persons?\s+employ/i.test(lower)) {
    return Math.max(12, Math.min(wideForm ? 18 : 20, Math.max(headerLen * 0.16, dataLen * 0.45, 12)));
  }

  // Form I TN subsistence: keep "date of payment" visible under the amount header.
  if (
    /amount/.test(lower) &&
    /allowance/.test(lower) &&
    /paid/.test(lower) &&
    /date/.test(lower) &&
    /payment/.test(lower)
  ) {
    return Math.max(14, Math.min(wideForm ? 20 : 22, headerLen * 0.22 + Math.min(dataLen, 12) * 0.35));
  }

  // Form 14 RJ young-person / hours / overtime columns — keep readable in landscape.
  if (/young\s+person|total\s+hours\s+worked|days\s+on\s+which\s+overtime|extent\s+of\s+overtime/i.test(lower)) {
    return Math.max(10, Math.min(16, headerLen * 0.18 + Math.min(dataLen, 12) * 0.3));
  }

  // Long Form 11-style narrative headers
  if (headerLen >= 55) {
    return Math.max(12, Math.min(wideForm ? 20 : 22, headerLen * 0.2 + Math.min(dataLen, 10) * 0.25));
  }
  if (headerLen >= 28) {
    return Math.max(8.5, Math.min(16, headerLen * 0.2 + Math.min(dataLen, 14) * 0.3));
  }

  const fromHeader = Math.min(Math.max(headerLen * 0.22, 4.5), wideForm ? 12 : 14);
  const fromData = Math.min(Math.max(dataLen * 0.45, 3.5), wideForm ? 11 : 14);
  return Math.max(fromHeader, fromData, 4.5);
};

/**
 * Raise columns that fell below a header-based minimum, then renormalize to usableWidth.
 */
const enforcePdfColumnMinWidths = (widths, headers, usableWidth, options = {}) => {
  const src = Array.isArray(widths) ? widths.map((w) => Math.max(0, Number(w) || 0)) : [];
  if (!src.length) return src;
  const total = Math.max(1, Number(usableWidth) || src.reduce((a, b) => a + b, 0));
  const compactSerial = options?.compactSerial === true;
  const tightSerial = options?.tightSerial === true;
  const compactSex = options?.compactSex === true;
  const mins = src.map((_, i) => {
    const h = String(headers?.[i] || '')
      .replace(/\s+/g, ' ')
      .trim();
    const lower = h.toLowerCase();
    if (isPdfSerialNumberHeader(h)) {
      return tightSerial
        ? Math.max(24, total * 0.02)
        : compactSerial
          ? Math.max(32, total * 0.045)
          : Math.max(52, total * 0.08);
    }
    if (/^(sex|gender|age|photo)$/i.test(h)) {
      return compactSex ? Math.max(18, total * 0.025) : Math.max(28, total * 0.045);
    }
    // Form XXVI: address / identity columns need a large floor; day leaves stay tiny.
    if (options?.formXXVIWordSafe === true && h && !/^\(?\s*\d{1,2}\s*\)?$/.test(h)) {
      const longest = formXXVITamilNaduLongestHeaderWord(h);
      if (/permanent/.test(lower) && /address/.test(lower)) {
        return Math.max(78, total * 0.1);
      }
      if (/local\s+address/.test(lower)) {
        return Math.max(70, total * 0.09);
      }
      if (/name of(?:\s+the)?\s+work/.test(lower)) {
        return Math.max(56, total * 0.07);
      }
      // Date of Termination / Signature of workman / Signature of contractor.
      if (isFormXXVITamilNaduTrailingWideHeader(h)) {
        return Math.max(72, total * 0.085);
      }
      if (longest.length >= 4) {
        const wordMin = Math.min(total * 0.1, Math.max(48, longest.length * 4.8));
        if (/designation|number of days|rate of wages|father|husband/i.test(lower)) {
          return Math.max(wordMin, total * 0.065);
        }
        return Math.max(wordMin * 0.9, total * 0.045);
      }
    }
    if (h.length >= 55) return Math.max(56, total * 0.09);
    if (
      /amount/.test(lower) &&
      /allowance/.test(lower) &&
      /paid/.test(lower) &&
      /date/.test(lower) &&
      /payment/.test(lower)
    ) {
      return Math.max(64, total * 0.1);
    }
    if (h.length >= 28 || /name|address|witness|signature|occupation|department/i.test(lower)) {
      return Math.max(36, total * 0.06);
    }
    return Math.max(18, total * 0.03);
  });

  let next = src.map((w, i) => Math.max(w, mins[i]));
  let sum = next.reduce((a, b) => a + b, 0) || 1;
  if (sum > total) {
    let overflow = sum - total;
    for (let pass = 0; pass < 8 && overflow > 0.01; pass += 1) {
      const shrinkable = next.map((w, i) => Math.max(0, w - mins[i]));
      const shrinkSum = shrinkable.reduce((a, b) => a + b, 0);
      if (shrinkSum < 0.01) {
        // All columns already at mins — scale as a last resort.
        next = next.map((w) => (w / (sum || 1)) * total);
        sum = total;
        overflow = 0;
        break;
      }
      const take = Math.min(overflow, shrinkSum);
      next = next.map((w, i) => w - (take * shrinkable[i]) / shrinkSum);
      sum = next.reduce((a, b) => a + b, 0) || 1;
      overflow = sum - total;
    }
  } else if (Math.abs(sum - total) > 0.01) {
    const extra = total - sum;
    const growable = next.map((w, i) => Math.max(0, w - mins[i]));
    const growSum = growable.reduce((a, b) => a + b, 0);
    if (growSum > 0.01) {
      next = next.map((w, i) => w + (extra * growable[i]) / growSum);
    } else {
      next = next.map((w) => (w / sum) * total);
    }
  }
  // Final normalize only when needed — avoid FP shrink below hard mins.
  const finalSum = next.reduce((a, b) => a + b, 0) || 1;
  if (Math.abs(finalSum - total) <= 0.05) {
    return next;
  }
  return next.map((w) => (w / finalSum) * total);
};

/** Leaf-header-aware weights so Form W amount columns stay wide enough for 6–7 digit figures. */
const formWTamilNaduColumnWeight = (headerText, maxDataLen) => {
  const h = String(headerText || '')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
  if (!h) return Math.max(6, Math.min(maxDataLen || 4, 9));
  if (/^(?:s|sr|si)\.?\s*no\b/.test(h)) return 3.2;
  if (/name of(?:\s+the)?\s+employee/.test(h)) return 13;
  if (/employee identification|identification\s*no/.test(h)) return 7.5;
  if (/number of days worked|days worked/.test(h)) return 4.5;
  if (/date of payment/.test(h)) return 7.5;
  if (/remarks/.test(h)) return 7;
  if (/receipt by employee|bank transaction/.test(h)) return 8;
  if (/unpaid accumulation|subsistence allowance/.test(h)) return 7;
  if (
    /basic wage|dearness|house rent|other allowance|overtime|gross wages|net wages|total deductions|provident fund|state insurance|labour welfare|advance|damages|pending recovery|any other deduction/.test(
      h
    )
  ) {
    return Math.max(7.2, Math.min(Math.max(maxDataLen || 0, 6) + 1.5, 10));
  }
  return Math.max(6, Math.min(maxDataLen || 4, 9));
};

/** Form XXVII TN Register of Wages — keep identity cols readable; amount cols compact. */
const formXXVIITamilNaduColumnWeight = (headerText, maxDataLen) => {
  const h = String(headerText || '')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
  if (!h) return Math.max(4.5, Math.min(maxDataLen || 3, 7));
  if (/^(?:s|sr|si|sl)\.?\s*no\b|^serial\s*(?:no|number)/.test(h)) return 3.2;
  if (/name of(?:\s+the)?\s+(?:work(?:man|er)|employee)/.test(h)) return 11;
  if (/father|husband/.test(h)) return 8;
  if (/designation|employee\s*(?:number|id|code)/.test(h)) return 6.5;
  if (/signature|thumb\s+impression|cheque/.test(h)) return 8;
  if (/unpaid|umpaid/.test(h)) return 6.5;
  if (/wage\s*period|number\s+worked|units\s+of\s+work|days\s+worked/.test(h)) return 5;
  if (/daily\s+rated|piece\s+rate|overtime\s+rate/.test(h)) return 5.5;
  if (
    /basic\s+wage|dearness|wash\s+allow|\bhra\b|\bstb\b|cash\s+in\s+lieu|ecca|gross\s+wages?|net\s+wages?|providen|esi|fines?|other\s+deduction|total\s+deduction|\bpt\b|uniform/.test(
      h
    )
  ) {
    return Math.max(5.5, Math.min(Math.max(maxDataLen || 0, 5) + 1.2, 8.5));
  }
  return Math.max(5, Math.min(maxDataLen || 4, 7.5));
};

/** Longest word in a Form XXVI heading — used so PDF never mid-splits "Permanent" / "Wages". */
const formXXVITamilNaduLongestHeaderWord = (headerText) => {
  const words = String(headerText || '')
    .replace(/[()/\\|,;:]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .split(' ')
    .filter(Boolean);
  if (!words.length) return '';
  return words.reduce((best, w) => (w.length > best.length ? w : best), '');
};

/** Form XXVI trailing cols that need extra PDF width (long header labels). */
const isFormXXVITamilNaduTrailingWideHeader = (headerText) => {
  const lower = String(headerText || '')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
  if (!lower) return false;
  if (/termination/.test(lower)) return true;
  if (/signature/.test(lower) && /contractor|representative/.test(lower)) return true;
  if (
    (/signature|thumb|impression/.test(lower) && /workman|worker/.test(lower)) ||
    (/thumb\s+impression/.test(lower) && !/contractor/.test(lower))
  ) {
    return true;
  }
  return false;
};

/**
 * Form XXVI TN Register of Employment — identity / address cols get most of the page;
 * day leaves stay compact so addresses are readable.
 */
const formXXVITamilNaduColumnWeight = (headerText, maxDataLen = 0) => {
  const h = String(headerText || '')
    .replace(/\s+/g, ' ')
    .trim();
  const lower = h.toLowerCase();
  const longest = formXXVITamilNaduLongestHeaderWord(h);
  const wordFloor = Math.max(6, Math.min(14, (longest.length || 4) * 0.85));
  const dataBoost = Math.min(Math.max(Number(maxDataLen) || 0, 0), 48) * 0.28;

  if (!h) return Math.max(5, Math.min(maxDataLen || 3, 8));
  if (/^\(?\s*\d{1,2}\s*\)?$/.test(h)) return 0.55;
  if (/^(?:s|sr|si|sl)\.?\s*no\b|^serial\s*(?:no|number)/.test(lower)) {
    return Math.max(4.2, wordFloor * 0.45);
  }
  if (/name of(?:\s+the)?\s+work/.test(lower)) return Math.max(14, wordFloor + dataBoost * 0.35);
  if (/age\s*(?:&|and)?\s*sex/.test(lower)) return Math.max(7, wordFloor);
  if (/permanent/.test(lower) && /address/.test(lower)) {
    return Math.max(22, wordFloor + 4 + dataBoost);
  }
  if (/local\s+address/.test(lower)) {
    return Math.max(20, wordFloor + 3 + dataBoost);
  }
  if (/designation|nature of work/.test(lower)) return Math.max(12, wordFloor + dataBoost * 0.25);
  if (/father|husband/.test(lower)) return Math.max(12, wordFloor + dataBoost * 0.2);
  if (/date of entry|entry into/.test(lower)) return Math.max(10, wordFloor);
  if (/rate of wages?/.test(lower)) return Math.max(9, wordFloor);
  if (/number of days|days of work|days worked/.test(lower)) return Math.max(10, wordFloor);
  // Date of Termination / Signature of workman / Signature of contractor — extra width.
  if (isFormXXVITamilNaduTrailingWideHeader(h)) {
    return Math.max(18, wordFloor + 4);
  }
  if (/daily\s+hours/.test(lower)) return 1.2;
  return Math.max(wordFloor, Math.min(dataBoost + 5, 14));
};

/**
 * Cap Form XXVI day columns and give freed width to address / identity columns.
 */
const redistributeFormXXVIPdfColumnWidths = (widths, headers, dayBand, usableWidth) => {
  const next = Array.isArray(widths) ? widths.map((w) => Math.max(0, Number(w) || 0)) : [];
  if (!next.length || !dayBand) return next;
  const total = Math.max(1, Number(usableWidth) || next.reduce((a, b) => a + b, 0));
  const dayStart = Number(dayBand.start);
  const dayEnd = Number(dayBand.end);
  if (!Number.isFinite(dayStart) || !Number.isFinite(dayEnd) || dayEnd < dayStart) return next;

  // ~1.4% of page each — enough for P/WO/CL, not address-eating.
  const dayCap = Math.max(9, total * 0.014);
  let freed = 0;
  for (let c = dayStart; c <= dayEnd && c < next.length; c += 1) {
    if (next[c] > dayCap) {
      freed += next[c] - dayCap;
      next[c] = dayCap;
    }
  }
  if (freed < 0.5) return next;

  const priority = (header) => {
    const lower = String(header || '')
      .replace(/\s+/g, ' ')
      .trim()
      .toLowerCase();
    // Prefer the three trailing data columns the user asked to widen.
    if (isFormXXVITamilNaduTrailingWideHeader(header)) return 6.5;
    if (/permanent/.test(lower) && /address/.test(lower)) return 3.2;
    if (/local\s+address/.test(lower)) return 3.0;
    if (/name of(?:\s+the)?\s+work/.test(lower)) return 2.4;
    if (/designation|father|husband/.test(lower)) return 1.8;
    if (/date of entry|rate of wages|number of days|age\s*(?:&|and)?\s*sex/.test(lower)) return 1.5;
    if (/^(?:s|sr|si|sl)\.?\s*no\b|^serial/.test(lower)) return 1.0;
    return 0.8;
  };

  const shareIdx = [];
  const shareW = [];
  for (let c = 0; c < next.length; c += 1) {
    if (c >= dayStart && c <= dayEnd) continue;
    const h = headers?.[c] || '';
    if (/^\(?\s*\d{1,2}\s*\)?$/.test(String(h).trim())) continue;
    shareIdx.push(c);
    shareW.push(priority(h));
  }
  const shareSum = shareW.reduce((a, b) => a + b, 0) || 1;
  shareIdx.forEach((c, i) => {
    next[c] += (shareW[i] / shareSum) * freed;
  });

  // Keep total exactly usableWidth.
  const sum = next.reduce((a, b) => a + b, 0) || 1;
  if (Math.abs(sum - total) > 0.5) {
    const scale = total / sum;
    for (let i = 0; i < next.length; i += 1) next[i] *= scale;
  }
  return next;
};

/**
 * Wrap PDF cell text on whole words only (never mid-word like "Permanen" / "t").
 * Falls back to one word per line when the column is narrow but still fits the longest word.
 */
const wrapPdfTextPreferWholeWords = (doc, text, maxWidth, lineCap = 12) => {
  const raw = String(text || '')
    .replace(/\s+/g, ' ')
    .trim();
  if (!raw) return [' '];
  const width = Math.max(Number(maxWidth) || 0, 6);
  const words = raw.split(' ').filter(Boolean);
  if (!words.length) return [' '];

  const lines = [];
  let current = '';
  for (let i = 0; i < words.length; i += 1) {
    const word = words[i];
    const trial = current ? `${current} ${word}` : word;
    let trialW = width + 1;
    try {
      trialW = doc.getTextWidth(trial);
    } catch (_) {
      trialW = trial.length * 4.2;
    }
    if (!current || trialW <= width) {
      current = trial;
      continue;
    }
    lines.push(current);
    current = word;
    if (lines.length >= lineCap) break;
  }
  if (current && lines.length < lineCap) lines.push(current);
  return lines.length ? lines.slice(0, lineCap) : [' '];
};

const resolveLeafHeaderTexts = (rows, tableStart, headerBandEnd, colCount) => {
  const headers = Array.from({ length: colCount }, () => '');
  for (let c = 0; c < colCount; c += 1) {
    let best = '';
    let numberFallback = '';
    for (let r = headerBandEnd; r >= tableStart; r -= 1) {
      const t = String(rows[r]?.[c] || '')
        .replace(/\s+/g, ' ')
        .trim();
      if (!t) continue;
      // Column index row "(9)" / "9" — keep looking upward for the real label.
      if (/^\(?\s*\d{1,2}\s*\)?$/.test(t)) {
        if (!numberFallback) numberFallback = t;
        continue;
      }
      if (isFormPGJDateOfMonthGroupLabel(t)) continue;
      if (isFormQMaharashtraParentGroupLabel(t)) {
        if (!best) best = t;
        continue;
      }
      // Prefer the deepest (leaf) non-group label; skip ultra-wide group banners alone.
      if (
        /^deductions$/i.test(t) ||
        /^advances$/i.test(t) ||
        /^damages\s*\/\s*fine$/i.test(t) ||
        /^details of injury$/i.test(t) ||
        isFormBTamilNaduAmountsDeductedGroupLabel(t)
      ) {
        if (!best) best = t;
        continue;
      }
      best = t;
      break;
    }
    headers[c] = best || numberFallback;
  }
  return headers;
};

const stripFieldLabelPrefix = (text, labelRe) =>
  String(text || '')
    .replace(labelRe, '')
    .replace(/^[\s.:\-–—]+/, '')
    .trim();

/**
 * Tamil Nadu CLRA Form XXVII — Register of Wages [Rule 78(1)(a)].
 * Not quarterly returns (Form XXVII AP / other states).
 */
const looksLikeFormXXVIITamilNaduRegisterPdfContext = (
  metaLines,
  rows,
  sheetName = '',
  fileName = ''
) => {
  const blob = [
    ...(metaLines || []),
    ...(rows || []).slice(0, 14).flat(),
    sheetName || '',
    fileName || ''
  ]
    .join(' ')
    .toLowerCase();
  if (!blob.trim()) return false;
  if (/quarterly\s+return|form[\s._-]*xxvii[\s._-]*quarter|form[\s._-]*27[\s._-]*quarter/.test(blob)) {
    return false;
  }
  const isXxvii =
    /form\s*xxvii\b/.test(blob) ||
    /form[\s._-]*xxvii/.test(blob) ||
    /form_xxvii/.test(blob) ||
    (/form[\s._-]*27\b/.test(blob) && /tamil/.test(blob));
  const isWagesRegister =
    /register\s+of\s+wages/.test(blob) ||
    /rule\s*78\s*\(\s*1\s*\)\s*\(\s*a\s*\)/.test(blob) ||
    (/basic\s+wage/.test(blob) &&
      /gross\s+wages?/.test(blob) &&
      /net\s+wages?/.test(blob) &&
      /deductions?/.test(blob));
  return isXxvii && isWagesRegister;
};

/** Pull month name from Form XXVII "Wage Period : May" banner (not body WAGE PERIOD column). */
const extractMonthFromWagePeriodLine = (line) => {
  const t = String(line || '').replace(/\s+/g, ' ').trim();
  if (!t) return '';
  // Skip Form W style "Wage Period from 1st May … to …"
  if (/wage\s*period\s+from\b/i.test(t)) return '';
  const m = t.match(
    /^wage\s*period\s*:?\s*(january|february|march|april|may|june|july|august|september|october|november|december)\b/i
  );
  if (m?.[1]) return m[1];
  return '';
};

/** Detect Form XXVI CLRA muster from sheet meta / header band. */
const detectFormXXVIPdfLayout = (metaLines, rows, tableStart) => {
  const blob = [...(metaLines || []), ...(rows || []).slice(0, Math.min(rows.length, 12)).flat()]
    .map((t) => String(t || ''))
    .join(' ')
    .toLowerCase();
  const isFormXXVI =
    /form\s*xxvi\b/.test(blob) ||
    (/register of employment of contractual?\s*labour/.test(blob) &&
      /tamil\s+nadu\s+contract\s+labour|see\s+rule\s*75/.test(blob));
  if (!isFormXXVI) return null;

  let principal = '';
  let contractor = '';
  let worksite = '';
  let month = '';
  let date = '';

  (metaLines || []).forEach((raw) => {
    const line = String(raw || '').replace(/\s+/g, ' ').trim();
    if (!line || isSystemGeneratedDocumentNote(line)) return;
    const lower = line.toLowerCase();
    if (/^form\s*xxvi\b/.test(lower)) return;
    if (/see\s+rule\s*75/.test(lower)) return;
    if (/register of employment of contractual?\s*labour/.test(lower)) return;

    if (/principal\s+employer/.test(lower)) {
      principal = stripFieldLabelPrefix(line, /name\s+and\s+address\s+of\s+(?:the\s+)?principal\s+employer/i);
      return;
    }
    if (/name\s+and\s+address\s+of\s+(?:the\s+)?contractor/.test(lower)) {
      contractor = stripFieldLabelPrefix(line, /name\s+and\s+address\s+of\s+(?:the\s+)?contractor\.?/i);
      return;
    }
    if (/nature\s+and\s+location\s+of\s+work|name\s+and\s+location\s+of\s+(?:the\s+)?work/.test(lower)) {
      worksite = stripFieldLabelPrefix(
        line,
        /(?:nature\s+and\s+location\s+of\s+work|name\s+and\s+location\s+of\s+(?:the\s+)?work\s*site?)\.?/i
      );
      return;
    }
    if (/^month\s*:?\s*/i.test(line)) {
      month = stripFieldLabelPrefix(line, /^month/i);
      return;
    }
    if (/^date\s*:?\s*/i.test(line) || /^year\s*:?\s*/i.test(line)) {
      date = stripFieldLabelPrefix(line, /^(?:date|year)/i);
    }
  });

  // Also scan early sheet rows for label/value pairs that meta split across columns.
  for (let r = 0; r < Math.min(tableStart || 0, rows.length); r += 1) {
    const row = rows[r] || [];
    for (let c = 0; c < row.length; c += 1) {
      const cell = String(row[c] || '').replace(/\s+/g, ' ').trim();
      if (!cell) continue;
      const lower = cell.toLowerCase();
      const next = String(row[c + 1] || '').replace(/\s+/g, ' ').trim();
      if (!principal && /principal\s+employer/.test(lower)) {
        principal = stripFieldLabelPrefix(cell, /name\s+and\s+address\s+of\s+(?:the\s+)?principal\s+employer/i) || next;
      }
      if (!contractor && /name\s+and\s+address\s+of\s+(?:the\s+)?contractor/.test(lower)) {
        contractor = stripFieldLabelPrefix(cell, /name\s+and\s+address\s+of\s+(?:the\s+)?contractor\.?/i) || next;
      }
      if (
        !worksite &&
        /nature\s+and\s+location\s+of\s+work|name\s+and\s+location\s+of\s+(?:the\s+)?work/.test(lower)
      ) {
        worksite =
          stripFieldLabelPrefix(
            cell,
            /(?:nature\s+and\s+location\s+of\s+work|name\s+and\s+location\s+of\s+(?:the\s+)?work\s*site?)\.?/i
          ) || next;
      }
      if (!month && /^month\s*:?\s*$/i.test(cell) && next && !/date|year|name and/i.test(next)) month = next;
      if (!date && /^(?:date|year)\s*:?\s*$/i.test(cell) && next && !/name and|nature/i.test(next)) date = next;
    }
  }

  return {
    title: 'FORM XXVI',
    reference: 'See Rule 75 of the Tamil Nadu Contract Labour (Regulation and Abolition) Rules, 1975',
    subtitle: 'Register of Employment of Contractual Labour',
    principal,
    contractor,
    worksite,
    month,
    date
  };
};

/**
 * Find the "Daily hours of work" day band (typically cols after Rate of Wages through day 31).
 */
const detectDailyHoursBand = (rows, tableStart, headerBandEnd, colCount) => {
  let labelRow = -1;
  let labelCol = -1;
  let labelText = 'Daily hours of work';

  for (let r = tableStart; r <= headerBandEnd && r < rows.length; r += 1) {
    for (let c = 0; c < colCount; c += 1) {
      const t = String(rows[r]?.[c] || '').trim();
      if (/daily\s+hours\s+of\s+work/i.test(t)) {
        labelRow = r;
        labelCol = c;
        labelText = t.replace(/\s+/g, ' ').trim();
        break;
      }
    }
    if (labelCol >= 0) break;
  }

  if (labelCol < 0) {
    // Fallback: Rate of Wages followed by a long run of day numbers 1..31
    let rateCol = -1;
    for (let r = tableStart; r <= headerBandEnd && r < rows.length; r += 1) {
      for (let c = 0; c < colCount; c += 1) {
        if (/rate\s+of\s+wages?/i.test(String(rows[r]?.[c] || ''))) rateCol = c;
      }
    }
    if (rateCol < 0 || rateCol + 20 >= colCount) return null;
    labelCol = rateCol + 1;
  }

  let trailingStart = colCount;
  for (let r = tableStart; r <= headerBandEnd && r < rows.length; r += 1) {
    for (let c = labelCol + 1; c < colCount; c += 1) {
      const t = String(rows[r]?.[c] || '');
      if (/number\s+of\s+days|signature|thumb|termination|contractor/i.test(t)) {
        trailingStart = Math.min(trailingStart, c);
      }
    }
  }

  // Prefer an explicit 1..N day-number run in the header band.
  let dayEnd = -1;
  for (let r = tableStart; r <= Math.min(headerBandEnd + 2, rows.length - 1); r += 1) {
    let run = 0;
    let end = -1;
    for (let c = labelCol; c < Math.min(trailingStart, colCount); c += 1) {
      const n = Number(String(rows[r]?.[c] || '').trim());
      if (Number.isInteger(n) && n >= 1 && n <= 31) {
        run += 1;
        end = c;
      } else if (run > 0 && !String(rows[r]?.[c] || '').trim()) {
        // allow blanks inside a collapsed merge
        continue;
      } else if (run > 0) {
        break;
      }
    }
    if (run >= 20) dayEnd = Math.max(dayEnd, end);
  }

  if (dayEnd < labelCol) {
    // Default Form XXVI: 31 day leaf columns
    dayEnd = Math.min(labelCol + 30, trailingStart - 1, colCount - 1);
  }
  if (dayEnd < labelCol) return null;

  return {
    labelRow,
    start: labelCol,
    end: dayEnd,
    label: labelText || 'Daily hours of work'
  };
};

const extractFieldValue = (label, value, { requireColon = true } = {}) => {
  const v = String(value || '').trim();
  if (!v) return requireColon ? `${label} :` : label;
  return `${label} : ${v}`;
};

/** Parse Form W Men / Women / young-person count box from meta + pre-table rows. */
const extractFormWGenderBox = (metaLines, rows, tableStart) => {
  const out = {
    total: '',
    men: '',
    women: '',
    maleYoung: '',
    femaleYoung: ''
  };
  let foundLabels = false;

  const assignCounts = (vals) => {
    if (!Array.isArray(vals) || vals.length < 4) return false;
    const nums = vals.map((v) => String(v ?? '').trim());
    if (!nums.every((n) => /^\d{1,5}$/.test(n))) return false;
    out.men = nums[0];
    out.women = nums[1];
    out.maleYoung = nums[2];
    out.femaleYoung = nums[3];
    return true;
  };

  const meta = (metaLines || [])
    .map((l) => String(l || '').replace(/\s+/g, ' ').trim())
    .filter(Boolean);

  meta.forEach((line) => {
    const m = line.match(/total\s+number\s+of\s+persons?\s+employed\s*:?\s*(\d{1,5})?/i);
    if (m && m[1] != null && m[1] !== '') out.total = m[1];
  });

  // Labels then counts in flat meta (skip Month/Year tokens between labels and numbers)
  for (let i = 0; i < meta.length; i += 1) {
    const a = meta[i].toLowerCase();
    const b = String(meta[i + 1] || '').toLowerCase();
    const c = String(meta[i + 2] || '').toLowerCase();
    const d = String(meta[i + 3] || '').toLowerCase();
    if (
      a === 'men' &&
      b === 'women' &&
      /male young person/.test(c) &&
      /female young person/.test(d)
    ) {
      foundLabels = true;
      const after = [];
      for (let j = i + 4; j < meta.length && after.length < 4; j += 1) {
        const tok = meta[j];
        if (/^\d{1,5}$/.test(tok)) after.push(tok);
        else if (/^(month|year)\b/i.test(tok)) continue;
        else if (after.length > 0) break;
      }
      assignCounts(after);
      break;
    }
  }

  // Dense sheet rows above the wage table
  const scanEnd = Math.min(rows?.length || 0, Math.max(tableStart || 0, 0) + 2, 20);
  for (let r = 0; r < scanEnd; r += 1) {
    const row = rows[r] || [];
    for (let c = 0; c < row.length; c += 1) {
      const cell = String(row[c] || '')
        .replace(/\s+/g, ' ')
        .trim();
      if (!cell) continue;
      const totalMatch = cell.match(/total\s+number\s+of\s+persons?\s+employed\s*:?\s*(\d{1,5})?/i);
      if (totalMatch && totalMatch[1]) out.total = totalMatch[1];

      if (/^men$/i.test(cell)) {
        const labels = [0, 1, 2, 3].map((k) =>
          String(row[c + k] || '')
            .replace(/\s+/g, ' ')
            .trim()
            .toLowerCase()
        );
        if (
          labels[0] === 'men' &&
          labels[1] === 'women' &&
          /male young person/.test(labels[2]) &&
          /female young person/.test(labels[3])
        ) {
          foundLabels = true;
          const next = rows[r + 1] || [];
          assignCounts([next[c], next[c + 1], next[c + 2], next[c + 3]]);
        }
      }
    }
  }

  if (!foundLabels && out.men === '' && out.women === '' && out.maleYoung === '' && out.femaleYoung === '') {
    return null;
  }

  // Always render the four boxes once labels (or counts) are present.
  if (out.men === '') out.men = '0';
  if (out.women === '') out.women = '0';
  if (out.maleYoung === '') out.maleYoung = '0';
  if (out.femaleYoung === '') out.femaleYoung = '0';

  if (out.total === '') {
    const sum =
      (Number(out.men) || 0) +
      (Number(out.women) || 0) +
      (Number(out.maleYoung) || 0) +
      (Number(out.femaleYoung) || 0);
    out.total = String(sum);
  }
  return out;
};

const normalizeMetaKey = (text) =>
  String(text || '')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();

const isStatutoryTitleMetaLine = (line) => {
  const t = String(line || '').replace(/\s+/g, ' ').trim();
  if (!t || isSystemGeneratedDocumentNote(t)) return false;
  const lower = t.toLowerCase();
  if (/^form\s*(?:no\.?\s*)?[-–.]?\s*[a-z0-9xivlc.]+\b/i.test(t) && t.length < 48) return true;
  if (/see\s+(?:sub-)?rule|\[see\s+|^\(see\s+|^\(?vide\s+rule/i.test(lower)) return true;
  if (/^\(?prescribed\s+under\b|^\[prescribed\s+under\b/i.test(t)) return true;
  if (/^register of\b/i.test(t)) return true;
  if (/^overtime\s+muster\s+roll\b|^muster\s+roll\b/i.test(t)) return true;
  if (/^list of\b/i.test(t) && t.length < 60) return true;
  if (/^wage\s+slip\b|^letter\s+of\b|^notice\s+of\b|^combined\b/i.test(t) && t.length < 80) {
    return true;
  }
  // Act / rules banner under the form name (Excel line 3–4)
  if (/^the\s+.+\b(act|rules)\b/i.test(t) && t.length <= 180 && !/:/.test(t)) return true;
  // License / return style short titles without a field colon
  if (
    !/:/.test(t) &&
    t.length <= 100 &&
    /form|register|return|notice|accident|wages|employment|muster|overtime/i.test(lower)
  ) {
    return true;
  }
  return false;
};

const isStatutoryFieldMetaLine = (line) => {
  const t = String(line || '').replace(/\s+/g, ' ').trim();
  if (!t || isSystemGeneratedDocumentNote(t)) return false;
  if (isStatutoryTitleMetaLine(t)) return false;
  const lower = t.toLowerCase();
  if (/^month\s*:?\s*/i.test(t) || /^date\s*:?\s*/i.test(t) || /^year\s*:?\s*/i.test(t)) return true;
  if (
    /name\s+and\s+address|nature\s+and\s+location|address\s+of|principal\s+employer|contractor|establishment|worksite|wage\s+period|from\s*to|location\s+of\s+work/i.test(
      lower
    )
  ) {
    return true;
  }
  // Form C GJ / Loan-Recoveries register — LIN rows are fields (not title bands).
  if (/labour\s+identification\s+no/i.test(lower)) return true;
  // Form A GJ header labels
  if (/^name\s+of\s+owner\b/i.test(t)) return true;
  if (/^name\s+of\s+establishment\b/i.test(t)) return true;
  // Generic "Label : value" administrative lines
  if (/^[^:]{3,80}:\s*\S/.test(t)) return true;
  return false;
};

/** Excel order for Form C Gujarat Loan/Recoveries header fields. */
const FORM_C_GJ_PDF_FIELD_RANK = (text) => {
  const t = String(text || '')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
  if (!t) return 99;
  if (/name\s+of\s+establishment/.test(t) && !/principal/.test(t)) return 0;
  if (/labour\s+identification\s+no/.test(t) && /principal\s+employer/.test(t)) return 3;
  if (/labour\s+identification\s+no/.test(t)) return 1;
  if (/principal\s+employer/.test(t)) return 2;
  return 50;
};

const looksLikeFormCGJLoanRecoveriesPdfContext = (metaLines, rows, sheetName = '') => {
  const blob = [...(metaLines || []), ...(rows || []).slice(0, 18).flat(), sheetName || '']
    .join(' ')
    .toLowerCase();
  return (
    (/\bform\s*c\b|form[\s._-]*c[\s._-]/.test(blob) || /form_c_gj|form[\s._-]*c[\s._-]*gj/.test(blob)) &&
    /register\s+of\s+(?:loan|recoveries|loan\s*\/\s*recoveries)/.test(blob)
  );
};

const reorderFormCGJPdfHeaderFields = (fields) => {
  const list = Array.isArray(fields) ? fields.filter((f) => String(f || '').trim()) : [];
  if (list.length < 2) return list;
  return [...list].sort((a, b) => FORM_C_GJ_PDF_FIELD_RANK(a) - FORM_C_GJ_PDF_FIELD_RANK(b));
};

const extractFormXXAPAdministrativeRows = (rows, tableStart, metaLines = []) => {
  const out = [
    [
      'Name and Address of Contractor. :',
      'Name and address of Establishment in/under which contract is carried on:'
    ],
    ['Nature and location of work. :', 'Name and address of Principal Employer :']
  ];
  const sourceRows = [
    ...rows.slice(0, Math.min(Number(tableStart) || 0, rows.length)),
    ...(metaLines || []).map((line) => [line])
  ];
  sourceRows.forEach((sourceRow) => {
    const cells = (sourceRow || [])
      .map((cell) => String(cell || '').replace(/\s+/g, ' ').trim())
      .filter(Boolean);
    if (!cells.length) return;
    const blob = cells.join(' ').toLowerCase();
    const target = /nature\s+and\s+location\s+of\s+work/.test(blob) || /principal\s+employer/.test(blob) ? 1 :
      /contractor/.test(blob) || /establishment/.test(blob) ? 0 : -1;
    if (target < 0) return;
    const leftIndex = cells.findIndex((cell) => /contractor|nature\s+and\s+location\s+of\s+work/i.test(cell));
    const rightIndex = cells.findIndex((cell) => /establishment|principal\s+employer/i.test(cell));
    const left = leftIndex >= 0 ? cells[leftIndex] : '';
    const right = rightIndex >= 0 ? cells[rightIndex] : '';
    const continuation = (index) =>
      index >= 0 && index + 1 < cells.length &&
      !/contractor|establishment|nature\s+and\s+location\s+of\s+work|principal\s+employer/i.test(cells[index + 1])
        ? ` ${cells[index + 1]}`
        : '';
    if (target === 0) {
      if (left) out[0][0] = left + continuation(leftIndex);
      if (right) out[0][1] = right + continuation(rightIndex);
    } else {
      if (left) out[1][0] = left + continuation(leftIndex);
      if (right) out[1][1] = right + continuation(rightIndex);
    }
  });
  return out;
};

const extractFormXVIIAPAdministrativeRows = (rows, tableStart, metaLines = []) => {
  const out = [
    ['', ''],
    ['', ''],
    ['', '']
  ];
  const cells = [
    ...rows.slice(0, Math.min(Number(tableStart) || 0, rows.length)).flat(),
    ...(metaLines || [])
  ]
    .map((cell) => String(cell || '').replace(/\s+/g, ' ').trim())
    .filter(Boolean);
  const fields = [
    { row: 0, side: 0, re: /contractor/i },
    // Tolerate Establishemnt typo and split "in/" / "under which contract" labels.
    { row: 0, side: 1, re: /establ(?:ishment|ishemnt)|under\s+which\s+contract/i },
    { row: 1, side: 0, re: /nature\s+and\s+location\s+of\s+work/i },
    { row: 1, side: 1, re: /principal\s+employer/i },
    { row: 2, side: 0, re: /wage\s+period/i }
  ];
  const labelRe =
    /contractor|establ(?:ishment|ishemnt)|under\s+which\s+contract|nature\s+and\s+location\s+of\s+work|principal\s+employer|wage\s+period/i;
  const isDuplicateToken = (value) => {
    const compact = String(value || '')
      .replace(/\s+/g, '')
      .toLowerCase();
    if (compact.length < 8) return false;
    const half = Math.floor(compact.length / 2);
    return compact.slice(0, half) === compact.slice(half, half * 2);
  };
  const dedupeValue = (value) => {
    let v = String(value || '').replace(/\s+/g, ' ').trim();
    if (!v) return '';
    // "AP-Tadipatri AP-Tadipatri" / "VAYONA ENERGY VAYONA ENERGY…"
    const parts = v.split(/\s{2,}|\s\|\s/);
    if (parts.length === 2 && parts[0].trim().toLowerCase() === parts[1].trim().toLowerCase()) {
      return parts[0].trim();
    }
    const tokens = v.split(/\s+/);
    if (tokens.length >= 4 && tokens.length % 2 === 0) {
      const mid = tokens.length / 2;
      const a = tokens.slice(0, mid).join(' ').toLowerCase();
      const b = tokens.slice(mid).join(' ').toLowerCase();
      if (a === b) return tokens.slice(0, mid).join(' ');
    }
    if (isDuplicateToken(v)) {
      return v.slice(0, Math.ceil(v.length / 2)).trim();
    }
    return v;
  };
  fields.forEach(({ row, side, re }) => {
    const index = cells.findIndex((cell) => re.test(cell));
    if (index < 0 || out[row][side]) return;
    const source = cells[index];
    const colon = source.indexOf(':');
    let value = colon >= 0 ? source.slice(colon + 1).trim() : '';
    if (!value && index + 1 < cells.length && !labelRe.test(cells[index + 1])) {
      value = cells[index + 1];
    }
    // Establishment value often sits on the next meta line (below the split label).
    if (!value && side === 1 && row === 0) {
      for (let i = index + 1; i < Math.min(index + 4, cells.length); i += 1) {
        if (labelRe.test(cells[i])) continue;
        value = cells[i];
        break;
      }
    }
    value = dedupeValue(value);
    const label = source.slice(0, colon >= 0 ? colon + 1 : source.length).trim();
    out[row][side] = value ? `${label} ${value}`.trim() : source;
  });
  return out;
};

const extractFormXVIAPHeaderFields = (rows, tableStart, metaLines = []) => {
  // Combine all text into one string - handle merged metaLines properly
  let fullText = [...(metaLines || []), ...rows.slice(0, Math.min(Number(tableStart) || 0, rows.length)).flat()]
    .map((v) => String(v || '').replace(/\s+/g, ' ').trim())
    .filter(Boolean)
    .join(' ');
 
  // Remove rule text from the beginning so it doesn't interfere with field extraction
  fullText = fullText.replace(/^.*?(?=\s+Name\s+and\s+Address\s+of\s+Contractor)/i, '');
 
  // Field patterns - with label text preserved for output
  const fieldPatterns = [
    {
      re: /name\s+and\s+address\s+of\s+(?:the\s+)?establishment[\s\S]{0,80}contract\s+is\s+carried\s+on\s*:/i,
      label: 'Name and Address of the Establishment in/ under which contract is carried on :',
      key: 'establishment'
    },
    { re: /address\s+of\s+the\s+establishment\s*:/i, label: 'Name and Address of the Establishment in/ under which contract is carried on :', key: 'establishment' },
    { re: /name\s+and\s+address\s+of\s+contractor\.?\s*:/i, label: 'Name and Address of Contractor. :', key: 'contractor' },
    { re: /name\s+and\s+address\s+of\s+principal\s+employer\.?\s*:/i, label: 'Name and address of Principal Employer :', key: 'employer' },
    { re: /nature\s+and\s+location\s+of\s+work\s*:/i, label: 'Nature and Location of work :', key: 'nature' },
    { re: /for\s+the\s+month\s+of\s*:/i, label: 'For the Month of :', key: 'month' }
  ];
 
  const extracted = {};
 
  // For each field, find its content by looking between field labels
  for (let i = 0; i < fieldPatterns.length; i++) {
    const current = fieldPatterns[i];
    const match = fullText.match(current.re);
   
    if (!match) continue;
   
    // Start position after the colon
    const labelEnd = match.index + match[0].length;
    let valueEnd = fullText.length;
   
    // Find where value ends: look for next field label
    const remaining = fullText.substr(labelEnd);
   
    // Check for next field label
    let nextFieldMatch = null;
    for (let j = i + 1; j < fieldPatterns.length; j++) {
      const nextMatch = remaining.match(fieldPatterns[j].re);
      if (nextMatch) {
        if (!nextFieldMatch || nextMatch.index < nextFieldMatch.index) {
          nextFieldMatch = nextMatch;
        }
        break;
      }
    }
   
    if (nextFieldMatch) {
      valueEnd = labelEnd + nextFieldMatch.index;
    }
   
    // Also stop at table header indicators (S.No, Name of the Employee, Dates, Remarks)
    const tableMatch = remaining.match(/\bS\.?\s*No\s*\b|\bName\s+of\s+the\s+Employee\b|\bDates\b|\bRemarks\b/i);
    if (tableMatch && tableMatch.index < valueEnd - labelEnd) {
      valueEnd = labelEnd + tableMatch.index;
    }
   
    // Extract and clean value
    let value = fullText.substring(labelEnd, valueEnd).trim();
    value = value.replace(/^[:\s]+|[\s]+$/g, '').trim();
    if (current.key === 'establishment') {
      value = value
        .replace(/\(?\s*vide\s+rule[\s\S]*?central\s*\/?\s*a\.?\s*p\.?\s*rules?\.?\)?/i, '')
        .replace(/[\s,;:-]+$/, '')
        .trim();
    }
    if (current.key === 'nature') {
      value = sanitizeFormXVIAPNatureValue(value);
    }
    if (current.key === 'month') {
      value = sanitizeFormXVIAPMonthValue(value);
    }
   
    // Only use if valid (not empty and not a form/rule indicator)
    if (value && !/^\s*(form|vide|rule|s\.?\s*no)/i.test(value) && value.length > 1) {
      if (!extracted[current.key]) extracted[current.key] = `${current.label} ${value}`;
    }
  }

  const stripFormXVIFieldValue = (line) => {
    const text = String(line || '');
    const colon = text.indexOf(':');
    return colon >= 0 ? text.slice(colon + 1).trim() : '';
  };
  const establishmentValue = stripFormXVIFieldValue(extracted.establishment);
  const employerValue = stripFormXVIFieldValue(extracted.employer);
  if (employerValue && (!establishmentValue || establishmentValue.length + 12 < employerValue.length)) {
    extracted.establishment =
      `Name and Address of the Establishment in/ under which contract is carried on : ${employerValue}`;
  } else if (establishmentValue) {
    extracted.establishment =
      `Name and Address of the Establishment in/ under which contract is carried on : ${establishmentValue}`;
  }
 
  return {
    fields: [
      extracted.establishment ||
        'Name and Address of the Establishment in/ under which contract is carried on :',
      extracted.contractor || 'Name and Address of Contractor. :',
      extracted.employer || 'Name and address of Principal Employer :'
    ],
    rightFields: [
      extracted.nature || 'Nature and Location of work :',
      extracted.month || 'For the Month of :',
      ''
    ]
  };
};

/**
 * Build bordered header model for any statutory form (titles + field rows).
 * Form XXVI keeps Excel-style Month/Date right column when available.
 * @param {object} [opts]
 * @param {string} [opts.preferredTitle] catalog form name (e.g. "Form 15 Part 1")
 * @param {string} [opts.fileName] draft / template file name
 */
const buildStatutoryPdfHeaderModel = (metaLines, rows, tableStart, sheetName = '', opts = {}) => {
  const preferredTitle = opts?.preferredTitle || '';
  const fileName = opts?.fileName || '';
  const preferForm15Part1 = looksLikeForm15Part1PreferredContext(
    preferredTitle,
    fileName,
    sheetName
  );
  const effectiveMetaLines = preferForm15Part1
    ? scrubForm15Part1MetaLines(metaLines)
    : metaLines;

  const isFormTKA = looksLikeFormTSEKarnatakaPdfContext(
    effectiveMetaLines,
    rows,
    sheetName,
    fileName
  );
  if (isFormTKA) {
    return {
      titles: getFormTSEKarnatakaHeaderTitles(effectiveMetaLines, rows, tableStart),
      fields: extractFormTSEKarnatakaHeaderFields(effectiveMetaLines, rows, tableStart),
      rightFields: [],
      formTKA: true,
      hasSystemNote: (effectiveMetaLines || []).some((l) => isSystemGeneratedDocumentNote(l)),
      isFormXXVI: false,
      isFormW: false,
      isFormXXVIIRegister: false,
      titleBoxFullBorder: false,
      hideRightBandSplit: false
    };
  }

  const isFormXIIIAP = looksLikeFormXIIIAPPdfContext(
    effectiveMetaLines,
    rows,
    sheetName,
    fileName
  );
  if (isFormXIIIAP) {
    return {
      titles: getFormXIIIAPHeaderTitles(effectiveMetaLines, rows, tableStart),
      fields: [],
      rightFields: [],
      formXIIIAP: true,
      formXIIIAPAdminRows: extractFormXIIIAPAdministrativeRows(
        effectiveMetaLines,
        rows,
        tableStart
      ),
      hasSystemNote: (effectiveMetaLines || []).some((l) => isSystemGeneratedDocumentNote(l)),
      isFormXXVI: false,
      isFormW: false,
      isFormXXVIIRegister: false,
      titleBoxFullBorder: false,
      hideRightBandSplit: false
    };
  }

  const isFormXXIIIAPSE = looksLikeFormXXIIIAPSEPdfContext(
    effectiveMetaLines,
    rows,
    sheetName
  );
  if (isFormXXIIIAPSE) {
    return {
      titles: getFormXXIIIAPSEHeaderTitles(effectiveMetaLines, rows, tableStart),
      fields: [],
      rightFields: [],
      formXXIIIAPSE: true,
      formXXIIIAPSEAdminLayout: extractFormXXIIIAPSEAdminLayout(
        effectiveMetaLines,
        rows,
        tableStart
      ),
      hasSystemNote: (effectiveMetaLines || []).some((l) => isSystemGeneratedDocumentNote(l)),
      isFormXXVI: false,
      isFormW: false,
      isFormXXVIIRegister: false,
      titleBoxFullBorder: false,
      hideRightBandSplit: false
    };
  }

  const isFormXXIIIGJOt = looksLikeFormXXIIIGJOtPdfContext(
    effectiveMetaLines,
    rows,
    sheetName,
    fileName
  );
  if (isFormXXIIIGJOt) {
    const titles = getFormXXIIIGJOtHeaderTitles(effectiveMetaLines, rows, tableStart);
    const fields = [];
    const seenGj = new Set(titles.map((t) => normalizeMetaKey(t)));
    (effectiveMetaLines || []).forEach((raw) => {
      expandStatutoryMetaSegments(raw).forEach((line) => {
        const text = String(line || '').replace(/\s+/g, ' ').trim();
        if (!text || isSystemGeneratedDocumentNote(text)) return;
        if (isStatutoryTitleMetaLine(text)) return;
        const key = normalizeMetaKey(text);
        if (seenGj.has(key)) return;
        if (
          /name and address|nature and location|principal employer|contractor|establishment/i.test(
            text
          )
        ) {
          seenGj.add(key);
          fields.push(text);
        }
      });
    });
    return {
      titles,
      fields,
      rightFields: [],
      formXXIIIGJOt: true,
      hasSystemNote: (effectiveMetaLines || []).some((l) => isSystemGeneratedDocumentNote(l)),
      isFormXXVI: false,
      isFormW: false,
      isFormXXVIIRegister: false,
      titleBoxFullBorder: true,
      hideRightBandSplit: true
    };
  }

  const isFormXXIAP = looksLikeFormXXIAPPdfContext(effectiveMetaLines, rows, sheetName);
  if (isFormXXIAP) {
    return {
      titles: getFormXXIAPHeaderTitles(effectiveMetaLines, rows),
      fields: [],
      rightFields: [],
      formXXIAP: true,
      formXXAPAdminRows: extractFormXXAPAdministrativeRows(rows, tableStart, effectiveMetaLines),
      hasSystemNote: (effectiveMetaLines || []).some((l) => isSystemGeneratedDocumentNote(l)),
      isFormXXVI: false,
      isFormW: false,
      isFormXXVIIRegister: false,
      titleBoxFullBorder: false,
      hideRightBandSplit: false
    };
  }

  const isFormXVIIAP = looksLikeFormXVIIAPPdfContext(effectiveMetaLines, rows, sheetName);
  if (isFormXVIIAP) {
    return {
      titles: getFormXVIIAPHeaderTitles(effectiveMetaLines, rows),
      fields: [],
      rightFields: [],
      formXVIIAP: true,
      formXVIIAPAdminRows: extractFormXVIIAPAdministrativeRows(rows, tableStart, effectiveMetaLines),
      hasSystemNote: (effectiveMetaLines || []).some((l) => isSystemGeneratedDocumentNote(l)),
      isFormXXVI: false,
      isFormW: false,
      isFormXXVIIRegister: false,
      titleBoxFullBorder: false,
      hideRightBandSplit: false
    };
  }
  const isFormXVIAP = looksLikeFormXVIAPPdfContext(effectiveMetaLines, rows, sheetName);
  if (isFormXVIAP) {
    const formXVIFields = extractFormXVIAPHeaderFields(rows, tableStart, effectiveMetaLines);
    return {
      titles: getFormXVIAPHeaderTitles(effectiveMetaLines, rows),
      ...formXVIFields,
      formXVIAP: true,
      hasSystemNote: (effectiveMetaLines || []).some((l) => isSystemGeneratedDocumentNote(l)),
      isFormXXVI: false,
      isFormW: false,
      isFormXXVIIRegister: false,
      titleBoxFullBorder: false,
      hideRightBandSplit: false
    };
  }

  const isFormXXAP = looksLikeFormXXAPPdfContext(effectiveMetaLines, rows, sheetName);
  if (isFormXXAP) {
    const titles = getFormXXAPHeaderTitles(effectiveMetaLines, rows);
    return {
      titles,
      fields: [],
      rightFields: [],
      formXXAP: true,
      formXXAPAdminRows: extractFormXXAPAdministrativeRows(
        rows,
        tableStart,
        effectiveMetaLines
      ),
      hasSystemNote: (effectiveMetaLines || []).some((l) => isSystemGeneratedDocumentNote(l)),
      isFormXXVI: false,
      isFormW: false,
      isFormXXVIIRegister: false,
      titleBoxFullBorder: false,
      hideRightBandSplit: false
    };
  }

  const xxvi = detectFormXXVIPdfLayout(effectiveMetaLines, rows, tableStart);
  if (xxvi) {
    return {
      titles: [xxvi.title, xxvi.reference, xxvi.subtitle].filter(Boolean),
      fields: [
        extractFieldValue('Name and address of Principal Employer', xxvi.principal),
        extractFieldValue('Name and Address of Contractor.', xxvi.contractor),
        extractFieldValue('Nature and location of work.', xxvi.worksite)
      ],
      rightFields: [
        extractFieldValue('Month', xxvi.month, { requireColon: false }),
        extractFieldValue('Date', xxvi.date, { requireColon: false }),
        ''
      ],
      hasSystemNote: (effectiveMetaLines || []).some((l) => isSystemGeneratedDocumentNote(l)),
      isFormXXVI: true,
      isFormW: false
    };
  }

  if (looksLikeFormXRajasthanEmploymentCardPdfContext(effectiveMetaLines, rows, sheetName)) {
    const employmentHeader = buildFormXRajasthanEmploymentHeaderModel(effectiveMetaLines);
    return {
      titles: employmentHeader.titles,
      fields: employmentHeader.fields,
      rightFields: [],
      genderBox: null,
      hasSystemNote: (effectiveMetaLines || []).some((l) => isSystemGeneratedDocumentNote(l)),
      isFormXXVI: false,
      isFormW: false,
      isFormXXVIIRegister: false,
      titleBoxFullBorder: false,
      hideRightBandSplit: false,
      formXRajasthanEmployment: true,
    };
  }

  if (looksLikeFormXIRajasthanServiceCertificatePdfContext(effectiveMetaLines, rows, sheetName)) {
    const serviceHeader = buildFormXIRajasthanServiceHeaderModel(effectiveMetaLines);
    return {
      titles: serviceHeader.titles,
      fields: serviceHeader.fields,
      rightFields: [],
      genderBox: null,
      hasSystemNote: (effectiveMetaLines || []).some((l) => isSystemGeneratedDocumentNote(l)),
      isFormXXVI: false,
      isFormW: false,
      isFormXXVIIRegister: false,
      titleBoxFullBorder: false,
      hideRightBandSplit: false,
    };
  }

  if (looksLikeFormXVRajasthanWageSlipPdfContext(effectiveMetaLines, rows, sheetName)) {
    const wageSlipHeader = buildFormXVRajasthanWageSlipHeaderModel(effectiveMetaLines);
    return {
      titles: wageSlipHeader.titles,
      fields: wageSlipHeader.fields,
      rightFields: [],
      genderBox: null,
      hasSystemNote: (effectiveMetaLines || []).some((l) => isSystemGeneratedDocumentNote(l)),
      isFormXXVI: false,
      isFormW: false,
      isFormXXVIIRegister: false,
      titleBoxFullBorder: true,
      hideRightBandSplit: false,
      formXVRajasthanWageSlip: true,
    };
  }

  if (looksLikeFormXVAPServiceCertificatePdfContext(effectiveMetaLines, rows, sheetName)) {
    const serviceHeader = buildFormXVAPServiceCertificateHeaderModel(
      effectiveMetaLines,
      rows,
      tableStart
    );
    return {
      titles: serviceHeader.titles,
      fields: serviceHeader.fields,
      rightFields: [],
      formXVAPService: true,
      formXVAPAdminRows: serviceHeader.adminRows,
      formXVAPFieldRows: serviceHeader.fieldRows,
      hasSystemNote: (effectiveMetaLines || []).some((l) => isSystemGeneratedDocumentNote(l)),
      isFormXXVI: false,
      isFormW: false,
      isFormXXVIIRegister: false,
      titleBoxFullBorder: false,
      hideRightBandSplit: false
    };
  }

  const isFormW = looksLikeFormWPdfContext(effectiveMetaLines, rows, tableStart);
  const isFormPGJ = looksLikeFormPGJGujaratPdfContext(
    effectiveMetaLines,
    rows,
    sheetName,
    fileName
  );
  const isFormQMaharashtra = looksLikeFormQMaharashtraPdfContext(
    effectiveMetaLines,
    rows,
    sheetName,
    fileName
  );
  const isFormXVIIITamilNadu = looksLikeFormXVIIITamilNaduPdfContext(
    effectiveMetaLines,
    rows,
    sheetName,
    fileName
  );
  const isFormBTamilNadu = looksLikeFormBTamilNaduPdfContext(
    effectiveMetaLines,
    rows,
    sheetName,
    fileName
  );
  const isFormXIXTamilNadu = looksLikeFormXIXTamilNaduPdfContext(
    effectiveMetaLines,
    rows,
    sheetName,
    fileName
  );
  const isFormXXVIIRegister = looksLikeFormXXVIITamilNaduRegisterPdfContext(
    effectiveMetaLines,
    rows,
    sheetName,
    fileName
  );
  const genderBox = isFormW ? extractFormWGenderBox(effectiveMetaLines, rows, tableStart) : null;

  const titles = [];
  const fields = [];
  const seen = new Set();
  let month = '';
  let date = '';
  let year = '';
  let hasSystemNote = false;

  const pushUnique = (list, text, { asTitle = false } = {}) => {
    const raw = String(text || '').replace(/\s+/g, ' ').trim();
    if (!raw) return;
    let key = normalizeMetaKey(raw);
    // Collapse "Form XXII" / "FORM XXII"
    if (asTitle && /^form\s+/i.test(raw)) {
      key = `form:${key.replace(/^form\s+/, '')}`;
    }
    if (seen.has(key)) return;
    seen.add(key);
    list.push(raw);
  };

  const isFormXIVEmploymentCardMeta = looksLikeFormXIVEmploymentCardPdfContext(
    effectiveMetaLines,
    rows,
    sheetName
  );
  // Form XIV GJ: Nature may include "GJ-Amreli\nJunior Engineer" — keep one field line.
  const headerMetaSource = isFormXIVEmploymentCardMeta
    ? (effectiveMetaLines || []).map((raw) =>
        String(raw || '')
          .replace(/\s*\r?\n+\s*/g, ' ')
          .replace(/\s+/g, ' ')
          .trim()
      )
    : effectiveMetaLines;
  const formAGJGujaratPdfContextHint = looksLikeFormAGJGujaratPdfContext(
    effectiveMetaLines,
    rows,
    sheetName
  );

  headerMetaSource.forEach((raw) => {
    expandStatutoryMetaSegments(raw).forEach((line) => {
      if (!line) return;
      if (formAGJGujaratPdfContextHint && isFormAGJGujaratExportTitleBandText(line)) return;
      if (preferForm15Part1 && isStandaloneFormXTitle(line)) return;
      if (isSystemGeneratedDocumentNote(line)) {
        hasSystemNote = true;
        return;
      }
      if (isFormBTamilNadu && isFormBTamilNaduPdfSignatoryText(line)) {
        return;
      }
      // Form C footnotes belong under the table (outside the box), not in the title band.
      if (isFormCWhereverApplicableFootnoteText(line) || isFormCDamageLossFineFootnoteText(line)) {
        return;
      }
      // Form D GJ footnotes belong under the muster-roll box, not in the title band.
      if (isFormDGJGujaratOuterFootnoteText(line)) {
        return;
      }
      // Form A GJ footnotes belong under the employee table, not in the title band.
      if (isFormAGJGujaratOuterFootnoteText(line)) {
        return;
      }
      // Form B GJ: Rate of Minimum Wages box is painted separately (Excel right-side model).
      if (isFormBGJMinimumWagesMetaText(line)) {
        return;
      }
      // Festival holiday boxes on Form 25 are numbered 1–5 — never print as title rows.
      if (/^\d{1,2}$/.test(line)) return;
      // Form 14 RJ column titles must never become full-width meta bands.
      if (
        /name\s+of\s+persons?\s+employ/i.test(line) ||
        /whether\s+young\s+person/i.test(line) ||
        /total\s+hours\s+worked\s+during/i.test(line) ||
        /days\s+on\s+which\s+overtime/i.test(line) ||
        /extent\s+of\s+overtime\s+worked/i.test(line)
      ) {
        return;
      }
      // Form 14 RJ footnotes are painted under the grid (above system note).
      if (isForm14RajasthanFootnoteText(line)) return;
      // Gender box labels/counts are painted as a dedicated bordered box — not as field lines
      if (/^(men|women|male young person|female young person)$/i.test(line)) return;
      if (/^\d{1,4}$/.test(line) && isFormW) return;
      // Total line is the top row of the gender box when counts exist
      if (genderBox && /total\s+number\s+of\s+persons?\s+employed/i.test(line)) {
        const m = line.match(/total\s+number\s+of\s+persons?\s+employed\s*:?\s*(\d{1,5})?/i);
        if (m?.[1] && !genderBox.total) genderBox.total = m[1];
        return;
      }

      if (!isFormPGJ && /^month\s*:?\s*/i.test(line) && !/name and|address|nature/i.test(line)) {
        month = stripFieldLabelPrefix(line, /^month/i) || month;
        return;
      }
      // Form XXVII banner: "Wage Period : May" → Month right field (not a left meta row)
      if (/^wage\s*period\s*:?\s*/i.test(line) && !/wage\s*period\s+from\b/i.test(line)) {
        const fromWp = extractMonthFromWagePeriodLine(line);
        if (fromWp) month = month || fromWp;
        else {
          const rest = stripFieldLabelPrefix(line, /^wage\s*period/i);
          if (rest && !/weekly|monthly|fn\b/i.test(rest)) month = month || rest;
        }
        return;
      }
      if (/^year\s*:?\s*/i.test(line) && !/name and|address|nature|entry|termination/i.test(line)) {
        year = stripFieldLabelPrefix(line, /^year/i);
        return;
      }
      if (/^(?:date)\s*:?\s*/i.test(line) && !/name and|address|nature|entry|termination|payment/i.test(line)) {
        date = stripFieldLabelPrefix(line, /^date/i);
        return;
      }
      if (isStatutoryTitleMetaLine(line)) {
        pushUnique(titles, line, { asTitle: true });
        return;
      }
      if (isStatutoryFieldMetaLine(line) || isFormMetaText(line)) {
        pushUnique(fields, line);
        return;
      }
      // Form XIV: designation fragments (e.g. "Junior Engineer" after a Nature newline) are not titles.
      if (
        isFormXIVEmploymentCardMeta &&
        !/:/.test(line) &&
        !/^form\s*xiv\b/i.test(line) &&
        !/^employment\s+card$/i.test(line) &&
        !/(?:see|vide)\s+rule\s*76/i.test(line)
      ) {
        return;
      }
      // Leftover meta — treat short lines as titles, longer as fields
      if (line.length <= 90 && !/:/.test(line)) pushUnique(titles, line, { asTitle: true });
      else pushUnique(fields, line);
    });
  });

  if (isFormW) {
    const hasFormWTitle = titles.some((t) => /form[\s._-]*w/i.test(t));
    if (!hasFormWTitle) titles.unshift('FORM W');
  }

  const sheet = String(sheetName || '').trim();
  // Never promote Excel sheet tab names (e.g. "LWF Act - Form C") into the PDF heading.
  // Titles already extracted from the sheet body are the printed form name.
  const hasFormTitle = titles.some((t) => /^form\s+/i.test(String(t || '')));
  if (
    sheet &&
    sheet !== 'Sheet1' &&
    titles.length === 0 &&
    !looksLikeExcelSheetTabName(sheet) &&
    !looksLikeExcelDraftFileLabel(sheet) &&
    !/^form\s*w$/i.test(sheet) &&
    !(hasFormTitle && /^form\s+/i.test(sheet)) &&
    !titles.some((t) => normalizeMetaKey(t).includes(normalizeMetaKey(sheet).slice(0, 12)))
  ) {
    titles.unshift(sheet);
  }

  const rightFields = [];
  if (
    !isFormPGJ &&
    (
    month !== '' ||
    date !== '' ||
    year !== '' ||
    /month|date|year|wage\s*period/i.test((effectiveMetaLines || []).join(' '))
    )
  ) {
    const hasMonthDateHint = (effectiveMetaLines || []).some((l) =>
      /^(month|date|year|wage\s*period)\b/i.test(String(l || '').trim())
    );
    if (hasMonthDateHint || isFormW || isFormXXVIIRegister) {
      rightFields.push(extractFieldValue('Month', month, { requireColon: false }));
      if (isFormW || year || isFormXXVIIRegister) {
        rightFields.push(extractFieldValue('Year', year || date, { requireColon: false }));
      } else {
        rightFields.push(extractFieldValue('Date', date, { requireColon: false }));
      }
    }
  }

  let finalTitles = titles.filter((t) => !looksLikeExcelDraftFileLabel(t));
  const alreadyForm15LeaveWithWages =
    finalTitles.some((t) => /form[\s._-]*15\b/i.test(String(t || ''))) &&
    finalTitles.some((t) => /register\s+of\s+leave\s+with\s+wages/i.test(String(t || '')));
  // Always drop bare FORM-X when this is Form 15 Part 1 (catalog) or the sheet already
  // shows FORM-15 + Leave with Wages (partial Excel rewrite left FORM-X behind).
  if (
    preferForm15Part1 ||
    alreadyForm15LeaveWithWages ||
    finalTitles.some((t) => isStandaloneFormXTitle(t)) &&
      finalTitles.some((t) => /form[\s._-]*15\b/i.test(String(t || '')))
  ) {
    finalTitles = rewriteForm15Part1PdfTitles(finalTitles);
  }
  // Last pass: never print Form X next to Form 15.
  finalTitles = finalTitles.filter((t) => !isStandaloneFormXTitle(t));

  let finalFields =
    fields.length === 0 && rightFields.length > 0
      ? rightFields.map(() => '')
      : fields;
  // Form C GJ: Name of Establishment, then Labour Identification No. (not under title band).
  if (looksLikeFormCGJLoanRecoveriesPdfContext(effectiveMetaLines, rows, sheetName)) {
    finalFields = reorderFormCGJPdfHeaderFields(finalFields);
    finalTitles = finalTitles.filter(
      (t) => !/labour\s+identification\s+no/i.test(String(t || ''))
    );
  }
  // Form K GJ: drop truncated intro under NOTICE; keep the full sentence above the table.
  if (
    looksLikeFormKGJGujaratPdfContext(
      effectiveMetaLines,
      rows,
      sheetName,
      fileName
    )
  ) {
    const rewritten = rewriteFormKGJGujaratPdfHeader(finalTitles, finalFields);
    finalTitles = rewritten.titles;
    finalFields = rewritten.fields;
  }
  // Form P GJ: Excel-model titles + Establishment / Employer / Month (no Form B min-wages box).
  if (isFormPGJ) {
    const rewritten = rewriteFormPGJGujaratPdfHeader(finalTitles, finalFields);
    finalTitles = rewritten.titles;
    finalFields = rewritten.fields;
  }
  // Form Q MH: FORM Q / See Rule 26(1) / MUSTER-ROLL + Establishment / Employer / Month.
  if (isFormQMaharashtra) {
    const rewritten = rewriteFormQMaharashtraPdfHeader(finalTitles, finalFields);
    finalTitles = rewritten.titles;
    finalFields = rewritten.fields;
  }
  // Form O GJ: Address → Authorized person → Notice → legal text → Details of Workers.
  if (
    looksLikeFormOGJGujaratPdfContext(
      effectiveMetaLines,
      rows,
      sheetName,
      fileName
    )
  ) {
    const rewritten = rewriteFormOGJGujaratPdfHeader(finalTitles, finalFields);
    finalTitles = rewritten.titles;
    finalFields = rewritten.fields;
  }

  let formAGJSplitRows = [];
  const isFormBGJ =
    looksLikeFormBGJGujaratPdfContext(effectiveMetaLines, rows, sheetName) ||
    looksLikeFormBRajasthanPdfContext(effectiveMetaLines, rows, sheetName);
  if (looksLikeFormAGJGujaratPdfContext(effectiveMetaLines, rows, sheetName)) {
    finalFields = reorderFormAGJPdfHeaderFields(finalFields);
    finalTitles = finalTitles.filter((t) => !isFormAGJGujaratExportTitleBandText(t));
  }
  if (isFormBGJ) {
    // Keep establishment / owner / LIN / principal / wage period; drop wage-rate box scraps.
    finalFields = finalFields.filter((f) => !isFormBGJMinimumWagesMetaText(f));
    finalTitles = finalTitles.filter(
      (t) =>
        !isFormBGJMinimumWagesMetaText(t) &&
        !/^(highly\s*skilled|skilled|semi[-\s]?skilled|un[-\s]?skilled)$/i.test(String(t || '').trim())
    );
    // Short labels without ":" can land in titles — move Form B admin lines into fields.
    const moved = [];
    finalTitles = finalTitles.filter((t) => {
      const line = String(t || '').trim();
      if (
        /name\s+of\s+(?:establishment|owner)|labour\s+identification|principal\s+employer|wage\s+period/i.test(
          line
        )
      ) {
        moved.push(line);
        return false;
      }
      return true;
    });
    if (moved.length) finalFields = [...finalFields, ...moved];
    // Harvest establishment bands still sitting in early table rows.
    const harvestUntil = Math.min(
      Array.isArray(rows) ? rows.length : 0,
      Math.max(0, Number(tableStart) || 0) + 12
    );
    for (let r = 0; r < harvestUntil; r += 1) {
      const row = rows[r] || [];
      const filled = row.map((c) => String(c || '').replace(/\s+/g, ' ').trim()).filter(Boolean);
      if (!filled.length) continue;
      const blob = filled.join(' ');
      if (isFormBGJWageRegisterColHeaderBlob(blob)) break;
      if (isFormBGJMinimumWagesMetaText(blob) || isFormBGJMinimumWagesBandRow(row)) continue;
      if (
        /name\s+of\s+(?:establishment|owner)|labour\s+identification|principal\s+employer|wage\s+period/i.test(
          blob
        )
      ) {
        filled.forEach((t) => {
          if (!isFormBGJMinimumWagesMetaText(t)) finalFields.push(t);
        });
      }
    }
    finalFields = reorderFormBGJPdfHeaderFields(finalFields);
  }

  if (isFormBTamilNadu) {
    const rewritten = rewriteFormBTamilNaduPdfHeader(finalTitles, finalFields);
    finalTitles = rewritten.titles;
    finalFields = rewritten.fields;
  }

  const isFormCTamilNaduLwf =
    (/form\s*-?\s*c\b/i.test(
      [...(effectiveMetaLines || []), sheetName || '', fileName || ''].join(' ')
    ) ||
      /lwf\b/i.test([...(effectiveMetaLines || []), sheetName || '', fileName || ''].join(' '))) &&
    (/unpaid\s+accumulations|details\s+of\s+fines|quarter\s+ending|register\s+of\s+fines/i.test(
      [...(effectiveMetaLines || []), ...(rows || []).slice(0, 8).flat()].join(' ')
    ) ||
      (Array.isArray(rows) &&
        rows.some((row) => isFormCLwfColHeaderBlob((row || []).join(' ')))));

  if (isFormCTamilNaduLwf) {
    const rewritten = rewriteFormCTamilNaduLwfPdfHeader(finalTitles, finalFields);
    finalTitles = rewritten.titles;
    finalFields = rewritten.fields;
  }

  if (isFormXVIIITamilNadu) {
    finalTitles = coalesceFormXVIIITamilNaduTitleLines(finalTitles);
  }

  const formBGJMinWagesBox = looksLikeFormBGJGujaratPdfContext(
    effectiveMetaLines,
    rows,
    sheetName
  )
    ? extractFormBGJMinimumWagesBox(effectiveMetaLines, rows, tableStart)
    : null;

  return {
    titles: finalTitles,
    fields: finalFields,
    rightFields,
    formAGJSplitRows,
    formAGJGujarat: looksLikeFormAGJGujaratPdfContext(effectiveMetaLines, rows, sheetName),
    formBGJGujarat: looksLikeFormBGJGujaratPdfContext(effectiveMetaLines, rows, sheetName),
    formLGJGujarat: looksLikeFormLGJGujaratPdfContext(
      effectiveMetaLines,
      rows,
      sheetName,
      fileName
    ),
    formPGJGujarat: isFormPGJ,
    formQMaharashtra: isFormQMaharashtra,
    formBGJMinWagesBox,
    genderBox,
    hasSystemNote,
    isFormXXVI: false,
    isFormW,
    isFormXXVIIRegister,
    formXVIIITamilNadu: isFormXVIIITamilNadu,
    formBTamilNadu: isFormBTamilNadu,
    formCTamilNaduLwf: isFormCTamilNaduLwf,
    formXIXTamilNadu: isFormXIXTamilNadu,
    /** Form XXVII: full title box, no Month/Year vertical divider. */
    titleBoxFullBorder: isFormXXVIIRegister,
    hideRightBandSplit: isFormXXVIIRegister
  };
};

/**
 * Excel-style bordered header for ALL statutory forms — vertical/horizontal line grid
 * for form name + administrative fields below.
 */
const uniqueExpandedTitleBands = (titles = []) => {
  const seen = new Set();
  const out = [];
  (Array.isArray(titles) ? titles : []).forEach((title) => {
    expandStatutoryMetaSegments(title).forEach((seg) => {
      const line = String(seg || '').replace(/\s+/g, ' ').trim();
      if (!line) return;
      const key = line.toLowerCase();
      if (seen.has(key)) return;
      seen.add(key);
      out.push(line);
    });
  });
  return out;
};

const paintBorderedStatutoryHeader = (doc, headerModel, layout, yStart) => {
  const { marginX, marginTop, marginBottom, usableWidth, pageHeight } = layout;
  let y = yStart;
  const x0 = marginX;
  const x1 = marginX + usableWidth;
  const padX = 6;
  const lineH = 10;
  const rawTitles =
    headerModel?.formXIIIAP || headerModel?.formXXIIIAPSE
      ? Array.isArray(headerModel.titles)
        ? headerModel.titles
        : []
      : Array.isArray(headerModel?.titles)
        ? headerModel.titles
        : [];
  const fields = Array.isArray(headerModel?.fields) ? headerModel.fields : [];
  const rightFields = Array.isArray(headerModel?.rightFields) ? headerModel.rightFields : [];
  const genderBox = headerModel?.genderBox || null;
  const formBGJMinWagesBox = headerModel?.formBGJMinWagesBox || null;
  const formXXAPAdminRows = Array.isArray(headerModel?.formXXAPAdminRows)
    ? headerModel.formXXAPAdminRows
    : [];
  const formXVIIAPAdminRows = Array.isArray(headerModel?.formXVIIAPAdminRows)
    ? headerModel.formXVIIAPAdminRows
    : [];
  const formXIIIAPAdminRows = Array.isArray(headerModel?.formXIIIAPAdminRows)
    ? headerModel.formXIIIAPAdminRows
    : [];
  const formXVAPAdminRows = Array.isArray(headerModel?.formXVAPAdminRows)
    ? headerModel.formXVAPAdminRows
    : [];
  const formXVAPFieldRows = Array.isArray(headerModel?.formXVAPFieldRows)
    ? headerModel.formXVAPFieldRows
    : [];
  const isApForm =
    headerModel?.formXXAP === true ||
    headerModel?.formXXIAP === true ||
    headerModel?.formXVIIAP === true ||
    headerModel?.formXVIAP === true ||
    headerModel?.formXIIIAP === true ||
    headerModel?.formXIXAP === true ||
    headerModel?.formXVAPService === true ||
    headerModel?.formXXIIIAPSE === true;
  const isFormXIXTamilNadu = headerModel?.formXIXTamilNadu === true;
  const titles =
    headerModel?.formXXIIIAPSE || headerModel?.formXIIIAP
      ? rawTitles
      : isApForm
        ? uniqueExpandedTitleBands(rawTitles)
        : rawTitles;
  const useRightBand = rightFields.some((t) => String(t || '').trim());
  const hideRightBandSplit = headerModel?.hideRightBandSplit === true;
  const titleBoxFullBorder = headerModel?.titleBoxFullBorder === true;
  const rightBandW = useRightBand ? Math.min(150, usableWidth * 0.22) : 0;
  const splitX = useRightBand ? x1 - rightBandW : x1;
  let paintedGenderBox = false;

  const ensureSpace = (h) => {
    if (y + h > pageHeight - marginBottom) {
      doc.addPage();
      y = marginTop;
    }
  };

  const strokeRect = (x, top, w, h) => {
    doc.setDrawColor(0, 0, 0);
    doc.setLineWidth(0.7);
    doc.rect(x, top, w, h, 'S');
  };

  const paintFullBand = (text, { bold = false, size = 9, align = 'center', minH = 16, singleLine = false } = {}) => {
    const line = String(text || '').trim();
    if (!line) return;
    const effectiveSize = isApForm ? 9 : size;
    doc.setFont('helvetica', bold ? 'bold' : 'normal');
    doc.setFontSize(effectiveSize);
    // Preserve Excel multiline cells (e.g. workman name + father's name).
    const wrapped = singleLine
      ? [line]
      : line
          .split(/\n+/)
          .map((part) => String(part || '').trim())
          .filter(Boolean)
          .flatMap((part) => doc.splitTextToSize(part, usableWidth - padX * 2));
    if (!wrapped.length) return;
    const h = Math.max(minH, wrapped.length * (effectiveSize + 2) + 8);
    ensureSpace(h);
    strokeRect(x0, y, usableWidth, h);
    doc.setTextColor(0, 0, 0);
    const textY = y + (h - wrapped.length * (effectiveSize + 2)) / 2 + effectiveSize;
    if (align === 'right') {
      doc.text(wrapped, x1 - padX, textY, { align: 'right' });
    } else if (align === 'left') {
      doc.text(wrapped, x0 + padX, textY);
    } else {
      doc.text(wrapped, (x0 + x1) / 2, textY, { align: 'center' });
    }
    y += h;
  };

  /**
   * Form XXVII: one outer box around stacked title lines (full L/R borders),
   * horizontal rules only between lines — no internal verticals.
   */
  const paintTitleBoxBlock = (titleSegs) => {
    const prepared = [];
    titleSegs.forEach((title) => {
      expandStatutoryMetaSegments(title).forEach((seg) => {
        const line = String(seg || '').trim();
        if (!line) return;
        const lower = line.toLowerCase();
        const isFormName = /^form\s+/i.test(line);
        const isRegister =
          /^register of\b|^overtime\s+muster\s+roll\b|^muster\s+roll\b|^list of\b|^wage\s+slip\b|^letter\s+of\b|^notice\s+of\b|^combined\b/i.test(
            line
          );
        const isRule = /see\s+(?:sub-)?rule|prescribed\s+under/i.test(lower);
        const isActBanner = /^the\s+.+\b(act|rules)\b/i.test(line);
        const bold = isFormName || isRegister || prepared.length === 0;
        const size = isApForm ? 9 : isFormName ? 11 : isRegister ? 10 : isRule || isActBanner ? 8 : 9;
        const minH = isFormName || isRegister ? 20 : 16;
        doc.setFont('helvetica', bold ? 'bold' : 'normal');
        doc.setFontSize(size);
        const wrapped = doc.splitTextToSize(line, usableWidth - padX * 2);
        const h = Math.max(minH, wrapped.length * (size + 2) + 8);
        prepared.push({ line, bold, size, wrapped, h });
      });
    });
    if (!prepared.length) return;

    const totalH = prepared.reduce((a, p) => a + p.h, 0);
    ensureSpace(totalH);
    const boxTop = y;
    strokeRect(x0, boxTop, usableWidth, totalH);
    doc.setDrawColor(0, 0, 0);
    doc.setLineWidth(0.7);
    // Reinforce outer verticals so they meet the table edge (finish at full width).
    doc.line(x0, boxTop, x0, boxTop + totalH);
    doc.line(x1, boxTop, x1, boxTop + totalH);

    prepared.forEach((p, idx) => {
      if (idx > 0) {
        doc.line(x0, y, x1, y);
      }
      doc.setFont('helvetica', p.bold ? 'bold' : 'normal');
      doc.setFontSize(p.size);
      doc.setTextColor(0, 0, 0);
      const textY = y + (p.h - p.wrapped.length * (p.size + 2)) / 2 + p.size;
      doc.text(p.wrapped, (x0 + x1) / 2, textY, { align: 'center' });
      y += p.h;
    });
  };

  const paintSplitBand = (leftText, rightText, { minH = 16, size = 8 } = {}) => {
    const effectiveSize = isApForm ? 9 : size;
    doc.setFont('helvetica', 'normal');
    doc.setFontSize(effectiveSize);
    const leftW = useRightBand ? splitX - x0 - padX * 2 : usableWidth - padX * 2;
    const leftWrapped = doc.splitTextToSize(String(leftText || ''), Math.max(leftW, 40));
    const rightWrapped =
      useRightBand && String(rightText || '').trim()
        ? doc.splitTextToSize(String(rightText || ''), rightBandW - padX * 2)
        : [];
    const h = Math.max(minH, Math.max(leftWrapped.length, rightWrapped.length || 1) * (effectiveSize + 2) + 6);
    ensureSpace(h);
    strokeRect(x0, y, usableWidth, h);
    // Form XXVII: keep Month/Year on the right without an internal vertical divider.
    if (useRightBand && !hideRightBandSplit) {
      doc.line(splitX, y, splitX, y + h);
    }
    doc.setTextColor(0, 0, 0);
    doc.text(leftWrapped, x0 + padX, y + 10);
    if (rightWrapped.length) {
      if (hideRightBandSplit) {
        doc.text(rightWrapped, x1 - padX, y + 10, { align: 'right' });
      } else {
        doc.text(rightWrapped, splitX + padX, y + 10);
      }
    }
    y += h;
  };

  /** Form W Excel-style Men / Women / young-person bordered count box. */
  const paintFormWGenderBox = () => {
    if (!genderBox || paintedGenderBox) return;
    paintedGenderBox = true;
    const labels = ['Men', 'Women', 'Male young person', 'Female young person'];
    const values = [
      genderBox.men !== '' ? genderBox.men : '0',
      genderBox.women !== '' ? genderBox.women : '0',
      genderBox.maleYoung !== '' ? genderBox.maleYoung : '0',
      genderBox.femaleYoung !== '' ? genderBox.femaleYoung : '0'
    ];
    const boxW = Math.min(460, Math.max(320, usableWidth * 0.42));
    // Excel places the box mid-sheet; keep it left of Month/Year when present.
    const boxX = useRightBand
      ? Math.max(x0, splitX - boxW - 8)
      : x0 + Math.max(0, (usableWidth - boxW) * 0.35);
    const totalH = 16;
    const labelH = 26;
    const valueH = 16;
    const h = totalH + labelH + valueH;
    ensureSpace(h + 4);

    const colW = boxW / 4;
    doc.setDrawColor(0, 0, 0);
    doc.setLineWidth(0.7);
    doc.setTextColor(0, 0, 0);

    // Total row (merged)
    doc.rect(boxX, y, boxW, totalH, 'S');
    doc.setFont('helvetica', 'normal');
    doc.setFontSize(8);
    const totalText = `Total number of persons employed: ${genderBox.total !== '' ? genderBox.total : values.reduce((a, b) => a + (Number(b) || 0), 0)}`;
    doc.text(totalText, boxX + 4, y + 11);

    // Label row
    const labelY = y + totalH;
    labels.forEach((label, i) => {
      const cx = boxX + i * colW;
      doc.rect(cx, labelY, colW, labelH, 'S');
      doc.setFont('helvetica', 'bold');
      doc.setFontSize(7);
      const wrapped = doc.splitTextToSize(label, Math.max(colW - 4, 20)).slice(0, 3);
      const textH = wrapped.length * 8;
      doc.text(wrapped, cx + colW / 2, labelY + (labelH - textH) / 2 + 7, { align: 'center' });
    });

    // Value row
    const valueY = labelY + labelH;
    values.forEach((val, i) => {
      const cx = boxX + i * colW;
      doc.rect(cx, valueY, colW, valueH, 'S');
      doc.setFont('helvetica', 'normal');
      doc.setFontSize(9);
      doc.text(String(val), cx + colW / 2, valueY + 11, { align: 'center' });
    });

    y += h + 6;
  };

  // One Excel wrapText title cell → one centered band per line (Form I / Form 10 style).
  if (titleBoxFullBorder && titles.length) {
    paintTitleBoxBlock(titles);
  } else {
    let titlePaintIdx = 0;
    titles.forEach((title) => {
      expandStatutoryMetaSegments(title).forEach((seg) => {
        const lower = String(seg).toLowerCase();
        const isFormName = /^form\s+/i.test(seg);
          const isRegister =
          /^register of\b|^overtime\s+muster\s+roll\b|^muster[\s-]*roll\b|^list of\b|^wage\s+slip\b|^letter\s+of\b|^notice\s+of\b|^combined\b|^format\s+of\s+wage\s+register\b/i.test(
            seg
          );
        const isXviiiMusterTitle = isFormXVIIITamilNaduWagesCumMusterTitle(seg);
        const isRule = /see\s+(?:sub-)?rule|prescribed\s+under/i.test(lower);
        const isActBanner = /^the\s+.+\b(act|rules)\b/i.test(seg);
        paintFullBand(seg, {
          bold: isFormXIXTamilNadu ? false : isFormName || isRegister || isXviiiMusterTitle || titlePaintIdx === 0,
          singleLine: isXviiiMusterTitle,
          size: isApForm
            ? 9
            : headerModel?.formXVIIITamilNadu
              ? isFormName || isXviiiMusterTitle
                ? 12
                : isRule || isActBanner
                  ? 9
                  : 11
            : headerModel?.formTKA
              ? isFormName
                ? 14
                : isRegister
                  ? 12
                  : isRule || isActBanner
                    ? 10
                    : 9
            : headerModel?.formBGJGujarat || headerModel?.formLGJGujarat || headerModel?.formPGJGujarat
              ? isFormName
                ? 13
                : isRegister
                  ? 12
                  : isRule || isActBanner
                    ? 10
                    : 11
              : isFormName
                ? 11
                : isRegister
                  ? 10
                  : isRule || isActBanner
                    ? 8
                    : 9,
          align: 'center',
          minH: headerModel?.formTKA
            ? isFormName || isRegister
              ? 26
              : 18
            : headerModel?.formBGJGujarat || headerModel?.formLGJGujarat || headerModel?.formPGJGujarat
              ? isFormName || isRegister
                ? 24
                : 18
              : isFormName || isRegister
                ? 20
                : 16
        });
        titlePaintIdx += 1;
      });
    });
  }

  if (headerModel?.formXXIIIAPSE && headerModel.formXXIIIAPSEAdminLayout) {
    const admin = headerModel.formXXIIIAPSEAdminLayout;
    const adminSize = 9;
    const adminLineH = adminSize + 3;
    const labelW = usableWidth * FORM_XXIII_APSE_ADMIN_WIDTHS.label;
    const valueW = usableWidth * 0.4;
    const restW = Math.max(usableWidth - labelW - valueW, 40);
    const rightCols = 5;
    const rightW = restW / rightCols;
    const joinParts = (...parts) => parts.map((p) => String(p || '').trim()).filter(Boolean).join(' ');
    const fromText = joinParts(admin.fromLabel || 'From:', admin.fromValue);
    const toText = joinParts(admin.toLabel || 'To:', admin.toValue);
    const row1Cells = [
      { text: admin.establishmentLabel || 'Name of Establishment / Shop:', width: labelW, bold: true },
      { text: admin.establishmentValue || '', width: valueW, bold: false },
      { text: admin.registrationLabel || 'Registration No.:', width: rightW, bold: true },
      { text: admin.registrationValue || '', width: rightW, bold: false },
      { text: admin.wagePeriodLabel || 'Wage Period', width: rightW, bold: true },
      { text: fromText, width: rightW, bold: false },
      { text: toText, width: rightW, bold: false }
    ];
    const row2Cells = [
      { text: admin.addressLabel || 'Address of the Establishment:', width: labelW, bold: true },
      { text: admin.addressValue || '', width: usableWidth - labelW, bold: false }
    ];

    const paintAdminCells = (cells) => {
      const wrapped = cells.map((cell) => {
        doc.setFont('helvetica', cell.bold ? 'bold' : 'normal');
        doc.setFontSize(adminSize);
        return {
          ...cell,
          lines: doc.splitTextToSize(String(cell.text || ''), Math.max(cell.width - padX * 2, 20))
        };
      });
      const h = Math.max(
        28,
        wrapped.reduce((max, cell) => Math.max(max, cell.lines.length), 1) * adminLineH + 10
      );
      ensureSpace(h);
      strokeRect(x0, y, usableWidth, h);
      doc.setDrawColor(0, 0, 0);
      doc.setLineWidth(0.7);
      let x = x0;
      wrapped.forEach((cell, idx) => {
        if (idx > 0) doc.line(x, y, x, y + h);
        doc.setFont('helvetica', cell.bold ? 'bold' : 'normal');
        doc.setFontSize(adminSize);
        doc.setTextColor(0, 0, 0);
        if (cell.lines.length) {
          doc.text(cell.lines, x + padX, y + 12);
        }
        x += cell.width;
      });
      y += h;
    };

    paintAdminCells(row1Cells);
    paintAdminCells(row2Cells);
  }

  if (
    (headerModel?.formXXAP || headerModel?.formXXIAP || headerModel?.formXVIIAP || headerModel?.formXIIIAP) &&
    (formXXAPAdminRows.length || formXVIIAPAdminRows.length || formXIIIAPAdminRows.length)
  ) {
    (formXIIIAPAdminRows.length
      ? formXIIIAPAdminRows
      : formXVIIAPAdminRows.length
        ? formXVIIAPAdminRows
        : formXXAPAdminRows
    ).forEach((row) => {
      const left = String(row[0] || '');
      const right = String(row[1] || '');
      doc.setFont('helvetica', 'bold');
      doc.setFontSize(isApForm ? 9 : headerModel?.formXIIIAP ? 9 : 8);
      const half = usableWidth / 2;
      const leftLines = doc.splitTextToSize(left, half - padX * 2);
      const rightLines = doc.splitTextToSize(right, half - padX * 2);
      const h = Math.max(
        headerModel?.formXVIIAP ? 52 : 30,
        Math.max(leftLines.length, rightLines.length || 1) * 10 + 10
      );
      ensureSpace(h);
      strokeRect(x0, y, usableWidth, h);
      doc.line(x0 + half, y, x0 + half, y + h);
      if (leftLines.length) doc.text(leftLines, x0 + padX, y + 16);
      if (rightLines.length) doc.text(rightLines, x0 + half + padX, y + 16);
      y += h;
    });
  }

  if (headerModel?.formXVAPService && formXVAPAdminRows.length) {
    formXVAPAdminRows.forEach((row) => {
      const left = String(row[0] || '');
      const right = String(row[1] || '');
      doc.setFont('helvetica', 'normal');
      doc.setFontSize(9);
      const half = usableWidth / 2;
      const leftLines = doc.splitTextToSize(left, half - padX * 2);
      const rightLines = doc.splitTextToSize(right, half - padX * 2);
      const h = Math.max(52, Math.max(leftLines.length, rightLines.length || 1) * 10 + 18);
      ensureSpace(h);
      strokeRect(x0, y, usableWidth, h);
      doc.line(x0 + half, y, x0 + half, y + h);
      if (leftLines.length) doc.text(leftLines, x0 + padX, y + 14);
      if (rightLines.length) doc.text(rightLines, x0 + half + padX, y + 14);
      y += h;
    });
  }

  // Form B GJ Excel model:
  // 1) Rate of Minimum Wages box (centered under titles)
  // 2) Establishment / owner / LIN / principal / wage period BELOW that box — outside any border
  if (headerModel?.formBGJGujarat) {
    const fieldSize = 11;
    const skills = Array.isArray(formBGJMinWagesBox?.skills)
      ? formBGJMinWagesBox.skills
      : ['Highly Skilled', 'Skilled', 'Semi-Skilled', 'Un-Skilled'];
    const dataRows = Array.isArray(formBGJMinWagesBox?.rows) ? formBGJMinWagesBox.rows : [];
    const mwTitle = String(formBGJMinWagesBox?.title || 'Rate of Minimum Wages').trim();
    const titleH = 22;
    const skillH = 24;
    const rowH = 18;
    const boxH = formBGJMinWagesBox ? titleH + skillH + Math.max(dataRows.length, 3) * rowH : 0;
    const boxW = Math.min(460, Math.max(300, usableWidth * 0.55));
    const boxX = x0 + Math.max(0, (usableWidth - boxW) / 2);

    // 1) Rate of Minimum Wages box centered under titles.
    if (formBGJMinWagesBox) {
      ensureSpace(boxH + 8);
      const boxY = y;
      doc.setDrawColor(0, 0, 0);
      doc.setLineWidth(0.7);
      doc.setTextColor(0, 0, 0);
      doc.rect(boxX, boxY, boxW, boxH, 'S');
      doc.setFont('helvetica', 'bold');
      doc.setFontSize(10);
      const titleLines = doc.splitTextToSize(mwTitle, boxW - 10).slice(0, 2);
      const titleTextH = titleLines.length * 11;
      doc.text(titleLines, boxX + boxW / 2, boxY + (titleH - titleTextH) / 2 + 10, {
        align: 'center',
      });
      doc.line(boxX, boxY + titleH, boxX + boxW, boxY + titleH);

      const labelW = boxW * 0.28;
      const skillW = (boxW - labelW) / 4;
      doc.setFont('helvetica', 'bold');
      doc.setFontSize(9);
      skills.forEach((skill, i) => {
        const cx = boxX + labelW + i * skillW;
        doc.line(cx, boxY + titleH, cx, boxY + boxH);
        const wrapped = doc.splitTextToSize(String(skill || ''), Math.max(skillW - 3, 18)).slice(0, 2);
        const skillTextH = wrapped.length * 10;
        doc.text(wrapped, cx + skillW / 2, boxY + titleH + (skillH - skillTextH) / 2 + 9, {
          align: 'center',
        });
      });
      doc.line(boxX + labelW, boxY + titleH, boxX + labelW, boxY + boxH);
      doc.line(boxX, boxY + titleH + skillH, boxX + boxW, boxY + titleH + skillH);

      const paintRows =
        dataRows.length >= 3
          ? dataRows.slice(0, 3)
          : [
              { label: 'Minimum Basic', values: ['', '', '', ''] },
              { label: 'DA', values: ['', '', '', ''] },
              { label: 'Overtime', values: ['', '', '', ''] },
            ];
      paintRows.forEach((row, ri) => {
        const ry = boxY + titleH + skillH + ri * rowH;
        if (ri > 0) doc.line(boxX, ry, boxX + boxW, ry);
        doc.setFont('helvetica', 'normal');
        doc.setFontSize(9);
        doc.text(String(row.label || ''), boxX + 4, ry + 12);
        (row.values || []).slice(0, 4).forEach((val, vi) => {
          const cx = boxX + labelW + vi * skillW;
          const text = String(val || '').trim();
          if (text) doc.text(text, cx + skillW / 2, ry + 12, { align: 'center' });
        });
      });
      y += boxH + 10;
    }

    // 2) Establishment fields BELOW the min-wages box — plain text, no border.
    doc.setFont('helvetica', 'normal');
    doc.setFontSize(fieldSize);
    fields.forEach((field) => {
      const line = String(field || '').trim();
      if (!line) return;
      const wrapped = doc.splitTextToSize(line, usableWidth - padX);
      if (!wrapped.length) return;
      const blockH = wrapped.length * (fieldSize + 2) + 3;
      ensureSpace(blockH);
      doc.setFont('helvetica', 'normal');
      doc.setFontSize(fieldSize);
      doc.setTextColor(0, 0, 0);
      doc.text(wrapped, x0, y + fieldSize);
      y += blockH;
    });
    y += 6;
  } else if (headerModel?.formXVAPService && formXVAPFieldRows.length) {
    formXVAPFieldRows.forEach(([left, right]) => paintSplitBand(left, right, { minH: 22, size: 9 }));
  } else {
    fields.forEach((field, idx) => {
    const right = rightFields[idx] || '';
    if (headerModel?.formXVAPService) {
      paintSplitBand(field, '', { minH: 22, size: 9 });
    } else if (useRightBand && idx < Math.max(rightFields.length, 2) && (right || idx < 2)) {
      paintSplitBand(field, right, {
        minH: 16,
        size: headerModel?.formXVIAP ? 9 : headerModel?.formAGJGujarat ? 10 : 8,
      });
    } else {
      // Form A Part-A label stays left-aligned (Excel + PDF); do not center it.
      paintFullBand(field, {
        bold: false,
        size: headerModel?.formTKA
          ? 11
          : headerModel?.formXVIIITamilNadu
            ? 10
          : headerModel?.formXVIAP
            ? 9
            : headerModel?.formAGJGujarat || headerModel?.formLGJGujarat || headerModel?.formXXIIIGJOt
              ? 10
              : 8,
        align: headerModel?.formLGJGujarat ? 'center' : 'left',
        minH: headerModel?.formTKA
          ? 22
          : headerModel?.formAGJGujarat || headerModel?.formLGJGujarat || headerModel?.formXXIIIGJOt
            ? 20
            : 16,
      });
    }
    // Place gender box after establishment / total band (Excel layout)
    if (
      genderBox &&
      !paintedGenderBox &&
      (/address of the establishment|total number of persons/i.test(String(field || '')) ||
        idx === 0)
    ) {
      // Prefer after establishment line; if first field is employer, still paint once early
      if (/address of the establishment|total number of persons/i.test(String(field || ''))) {
        paintFormWGenderBox();
      }
    }
  });
  }
  if (genderBox && !paintedGenderBox) {
    paintFormWGenderBox();
  }

  return y;
};

/**
 * Draw one sheet: meta headings once (full width), then full data table.
 */
const drawMatrixSheet = (doc, matrix, startY, pdfOpts = {}) => {
  if (matrix?.formMGJLayout === true) {
    return drawFormMGJGujaratIdentityCard(doc, matrix, startY);
  }
  if (matrix?.formNGJLayout === true) {
    return drawFormNGJGujaratLeaveBook(doc, matrix, startY);
  }
  if (matrix?.formFKALayout === true) {
    return drawFormFKarnatakaLeaveRegister(doc, matrix, startY);
  }
  if (matrix?.formPKALayout === true) {
    return drawFormPKarnatakaNotice(doc, matrix, startY);
  }
  const pageWidth = doc.internal.pageSize.getWidth();
  const pageHeight = doc.internal.pageSize.getHeight();
  const marginX = 24;
  const marginTop = 28;
  const marginBottom = 32;
  const usableWidth = pageWidth - marginX * 2;
  const colCount = Math.max(1, matrix.colCount);
  const rows = matrix.rows;
  const tableStart = Math.max(0, matrix.tableStartRow || 0);
  const metaLines = Array.isArray(matrix.metaLines) ? matrix.metaLines : [];
  const headerModel = buildStatutoryPdfHeaderModel(metaLines, rows, tableStart, matrix.name, {
    preferredTitle: pdfOpts.preferredTitle || '',
    fileName: pdfOpts.fileName || ''
  });
  if (matrix.formXIXAPLayout === true) headerModel.formXIXAP = true;
  const isFormXIXKA = matrix.formXIXKarnatakaLayout === true;
  const isFormXIXTamilNadu =
    headerModel.formXIXTamilNadu === true ||
    looksLikeFormXIXTamilNaduPdfContext(
      metaLines,
      rows,
      matrix.name || '',
      pdfOpts.fileName || matrix.fileName || ''
    );
  if (isFormXIXTamilNadu) headerModel.formXIXTamilNadu = true;
  if (isFormXIXKA) {
    const kaTitles = [];
    const kaFields = [];
    (metaLines || []).forEach((raw) => {
      const line = String(raw || '').replace(/\s+/g, ' ').trim();
      if (!line) return;
      if (/^form\s*xix\b/i.test(line) || /see\s+rule\s*78/i.test(line) || /^wages?\s+slip\b/i.test(line)) {
        kaTitles.push(line);
      } else {
        kaFields.push(line);
      }
    });
    headerModel.titles = kaTitles.length ? kaTitles : headerModel.titles;
    headerModel.fields = kaFields;
    headerModel.rightFields = [];
    headerModel.titleBoxFullBorder = false;
    headerModel.hideRightBandSplit = false;
  }
  const isFormXIVKALayout = matrix.formXIVKALayout === true;
  const isFormQKALayout = matrix.formQKALayout === true;
  if (isFormXIVKALayout) {
    headerModel.fields = (headerModel.fields || []).filter((field) => {
      const text = String(field || '');
      return matchFormXIVPdfHeaderFieldIndex(text) < 0 && !isFormXIVEmploymentCardHeaderLabelText(text);
    });
  }
  if (isFormQKALayout) {
    headerModel.fields = [];
    headerModel.rightFields = [];
  }
  const isFormXXAP = headerModel.formXXAP === true;
  const isFormXXIAP = headerModel.formXXIAP === true;
  const isFormXVIIAP = headerModel.formXVIIAP === true;
  const isFormXVIAP = headerModel.formXVIAP === true;
  const isFormXIIIAP = headerModel.formXIIIAP === true;
  const isFormXXIIIAPSE = headerModel.formXXIIIAPSE === true;
  const isFormXXIIIGJOt = headerModel.formXXIIIGJOt === true;
  const isFormXIXAP = matrix.formXIXAPLayout === true;

  let headerBandEnd = tableStart;
  for (let r = tableStart; r < Math.min(rows.length, tableStart + 8); r += 1) {
    if (isLikelyHeaderBandRow(rows[r], r)) headerBandEnd = r;
  }
  // Form XIV Employment Card: every workman row is label|value data — no header band.
  const isFormXIVEmploymentCardSheet = looksLikeFormXIVEmploymentCardPdfContext(
    metaLines,
    rows,
    matrix.name || ''
  );
  if (isFormXIVEmploymentCardSheet) {
    headerBandEnd = tableStart - 1;
  }
  if (isFormQKALayout) {
    headerBandEnd = tableStart - 1;
  }
  if (isFormXIXAP) {
    headerBandEnd = tableStart - 1;
  }
  if (isFormXIXTamilNadu) {
    // Entire sheet is label|value data — no bold header band / centered headers.
    headerBandEnd = tableStart - 1;
  }
  if (isFormXIXKA) {
    headerBandEnd = Math.min(rows.length - 1, tableStart + 1);
  }
  if (isFormXXAP) {
    headerBandEnd = Math.min(rows.length - 1, tableStart + 1);
  }
  if (isFormXXIIIAPSE) {
    for (let r = tableStart; r < Math.min(rows.length, tableStart + 4); r += 1) {
      if (isFormXXIIIAPSETableHeaderRow(rows[r])) {
        headerBandEnd = Math.max(headerBandEnd, r);
      }
    }
    const xxiiiBand = findFormXXIIIAPSEDeductionsBand(rows, tableStart, colCount);
    if (xxiiiBand) {
      headerBandEnd = Math.max(headerBandEnd, xxiiiBand.parentRow, xxiiiBand.childRow);
    }
  }
  if (isFormXVIIAP) {
    const formXVIIHeaderStart = Math.max(0, tableStart - 4);
    const formXVIIHeaderEnd = Math.min(rows.length - 1, tableStart + 3);
    for (let r = formXVIIHeaderStart; r <= formXVIIHeaderEnd; r += 1) {
      const blob = (rows[r] || []).join(' ').toLowerCase();
      if (
        /deductions\s*,?\s*if\s+any|wage\s+period|register\s+of\s+workmen|name\s+of\s+workmen/.test(
          blob
        )
      ) {
        headerBandEnd = Math.max(headerBandEnd, r);
      }
    }
  }
  if (isFormXVIAP) {
    for (let r = Math.max(0, tableStart - 4); r <= Math.min(rows.length - 1, tableStart + 4); r += 1) {
      const blob = (rows[r] || []).join(' ').toLowerCase();
      if (/(^|\s)dates?(\s|$)|attendance|muster\s+roll/.test(blob)) {
        headerBandEnd = Math.max(headerBandEnd, r);
      }
    }
  }
  // Form B TN LWF — parent "Amounts deducted during the month" + Fine / Other leaves.
  if (headerModel.formBTamilNadu) {
    for (let r = tableStart; r < Math.min(rows.length, tableStart + 6); r += 1) {
      const blob = (rows[r] || []).join(' ').toLowerCase();
      if (
        /total\s+number\s+of\s+empl|emolument|amounts?\s+deducted|actually\s+paid|balance\s+due|other\s+deductions?|\bfine\b/.test(
          blob
        )
      ) {
        headerBandEnd = Math.max(headerBandEnd, r);
      }
    }
  }
  // Form W has a 3-row merged header band (group → mid → leaf).
  if (headerModel.isFormW) {
    for (let r = tableStart; r < Math.min(rows.length, tableStart + 5); r += 1) {
      const blob = (rows[r] || []).join(' ').toLowerCase();
      if (/basic wage|dearness|house rent|deductions|net wages|s\.?\s*no/.test(blob)) {
        headerBandEnd = Math.max(headerBandEnd, r);
      }
    }
  }
  // Form T KA — group banners + day numbers + statutory column-index strip.
  if (headerModel.formTKA === true || matrix.formTKALayout === true) {
    for (let r = tableStart; r < Math.min(rows.length, tableStart + 6); r += 1) {
      const blob = (rows[r] || []).join(' ').toLowerCase();
      if (
        /s\.?\s*no|attendance|earned\s+wages|deductions|wages\s+fixed|name of(?:\s+the)?\s+employee|principal\s+employer|father/.test(
          blob
        )
      ) {
        headerBandEnd = Math.max(headerBandEnd, r);
      }
      const filled = (rows[r] || []).filter((c) => String(c || '').trim());
      if (filled.filter((c) => /^\d{1,2}$/.test(String(c).trim())).length >= 8) {
        headerBandEnd = Math.max(headerBandEnd, r);
      }
    }
  }
  // Form XXVII Register of Wages — group → mid → leaf → column-number rows.
  if (headerModel.isFormXXVIIRegister) {
    for (let r = tableStart; r < Math.min(rows.length, tableStart + 6); r += 1) {
      const blob = (rows[r] || []).join(' ').toLowerCase();
      if (
        /wages?\s+earned|deductions?|basic\s+wage|gross\s+wages?|net\s+wages?|other\s+allowances?|wash\s+allow/.test(
          blob
        )
      ) {
        headerBandEnd = Math.max(headerBandEnd, r);
      }
      const filled = (rows[r] || []).filter((c) => String(c || '').trim());
      if (filled.filter((c) => /^\d{1,2}$/.test(String(c).trim())).length >= 8) {
        headerBandEnd = Math.max(headerBandEnd, r);
      }
    }
  }
  // Leave registers + Form 15 Part 2 wage bands + Form VI festival band.
  // Form Q KA appointment order: "3. Name of the Employee" is a data row, not a header band.
  if (!isFormQKALayout) {
    for (let r = tableStart; r < Math.min(rows.length, tableStart + 8); r += 1) {
      const blob = (rows[r] || []).join(' ').toLowerCase();
      if (
        /earned\s+leave|medical\s+leave|other\s+leave|maternity\s+benefit|leave\s+at\s+the\s+beginning|leave\s+availed|name\s+of\s+the\s+employee/.test(
          blob
        ) ||
        /deductions|advances|damages\s*\/\s*fine|leave\s+wages|provident\s+fund|gross\s+wages|net\s+wages|advance\s+paid/.test(
          blob
        ) ||
        /days,?\s+dates\s+and\s+months|national\s+and\s+festival\s+holidays|pongal|republic\s+day|diwali|christmas/.test(
          blob
        )
      ) {
        headerBandEnd = Math.max(headerBandEnd, r);
      }
    }
  }
  // Form P GJ — group banner + day numbers 1–31 + statutory column-index strip.
  if (headerModel.formPGJGujarat === true) {
    for (let r = tableStart; r < Math.min(rows.length, tableStart + 6); r += 1) {
      const blob = (rows[r] || []).join(' ').toLowerCase();
      if (
        /sr\.?\s*no|date\s+of\s+(the\s+)?month|full\s+name\s+of\s+the\s+worker|interval\s+for\s+rest|working\s+hours/.test(
          blob
        )
      ) {
        headerBandEnd = Math.max(headerBandEnd, r);
      }
      const filled = (rows[r] || []).filter((c) => String(c || '').trim());
      if (filled.filter((c) => /^\(?\s*\d{1,2}\s*\)?$/.test(String(c).trim())).length >= 8) {
        headerBandEnd = Math.max(headerBandEnd, r);
      }
    }
  }
  // Form Q MH — same Excel header stack (parent groups + day numbers + col index).
  if (headerModel.formQMaharashtra === true) {
    for (let r = tableStart; r < Math.min(rows.length, tableStart + 6); r += 1) {
      const blob = (rows[r] || []).join(' ').toLowerCase();
      if (
        /sr\.?\s*no|date\s+of\s+(the\s+)?month|full\s+name\s+of\s+the\s+worker|interval\s+for\s+rest|working\s+hours|deductions?/.test(
          blob
        )
      ) {
        headerBandEnd = Math.max(headerBandEnd, r);
      }
      const filled = (rows[r] || []).filter((c) => String(c || '').trim());
      if (filled.filter((c) => /^\(?\s*\d{1,2}\s*\)?$/.test(String(c).trim())).length >= 8) {
        headerBandEnd = Math.max(headerBandEnd, r);
      }
    }
  }
  // Form Q KA: every label|value row is body text (Excel model) — never a header band.
  if (isFormQKALayout) {
    headerBandEnd = tableStart - 1;
  }
  const formXVIDateBand = isFormXVIAP
    ? (() => {
        for (let r = Math.max(0, tableStart - 4); r <= Math.min(rows.length - 1, tableStart + 5); r += 1) {
          const row = rows[r] || [];
          const dateStart = row.findIndex((cell) => /^1$/.test(String(cell || '').trim()));
          if (dateStart < 0) continue;
          let dateEnd = dateStart;
          for (let n = 1; n <= 31; n += 1) {
            if (String(row[dateStart + n - 1] || '').trim() !== String(n)) break;
            dateEnd = dateStart + n - 1;
          }
          if (dateEnd - dateStart + 1 >= 20) {
            const labelRow = Math.max(0, r - 1);
            return { labelRow, start: dateStart, end: Math.min(dateStart + 30, colCount - 1), label: 'Dates' };
          }
        }
        return null;
      })()
    : null;
  const isFormTKASheet = headerModel.formTKA === true || matrix.formTKALayout === true;
  const formTKAAttendanceBand = isFormTKASheet
    ? matrix.attendanceBand && Number.isFinite(matrix.attendanceBand.start)
      ? matrix.attendanceBand
      : detectFormTSEKarnatakaAttendanceBand(rows, tableStart, headerBandEnd, colCount)
    : null;
  const formLGJGroupBands = looksLikeFormLGJGujaratPdfContext(
    metaLines,
    rows,
    matrix.name || '',
    pdfOpts.fileName || matrix.fileName || ''
  )
    ? detectFormLGJGujaratPdfGroupBands(rows, tableStart, headerBandEnd, colCount)
    : [];
  const formPGJGroupBands = looksLikeFormPGJGujaratPdfContext(
    metaLines,
    rows,
    matrix.name || '',
    pdfOpts.fileName || matrix.fileName || ''
  )
    ? detectFormPGJGujaratPdfGroupBands(rows, tableStart, headerBandEnd, colCount)
    : [];
  const formQMhGroupBands = looksLikeFormQMaharashtraPdfContext(
    metaLines,
    rows,
    matrix.name || '',
    pdfOpts.fileName || matrix.fileName || ''
  )
    ? detectFormQMaharashtraPdfGroupBands(rows, tableStart, headerBandEnd, colCount)
    : [];
  const formQMhDateBand = formQMhGroupBands.find(
    (b) => b && isFormQMaharashtraDateOfMonthGroupLabel(b.label) && b.end > b.start
  );
  const dayBand =
    formXVIDateBand ||
    formTKAAttendanceBand ||
    (formPGJGroupBands[0] && formPGJGroupBands[0].end > formPGJGroupBands[0].start
      ? formPGJGroupBands[0]
      : null) ||
    formQMhDateBand ||
    detectDailyHoursBand(rows, tableStart, headerBandEnd, colCount);
  // Form Q MH: use Form-Q group bands only for Working hours / Interval / Date / Deductions
  // so generic "Deductions" detection cannot paint two split banners.
  const statutoryGroupBands = detectStatutoryGroupHeaderBands(
    rows,
    tableStart,
    headerBandEnd,
    colCount
  ).filter((band) => {
    if (!(formQMhGroupBands.length > 0)) return true;
    return !isFormQMaharashtraParentGroupLabel(band?.label);
  });
  const groupBands = [
    ...statutoryGroupBands,
    ...formLGJGroupBands,
    ...formPGJGroupBands,
    ...formQMhGroupBands,
  ];
  if (headerModel.formBTamilNadu) {
    detectFormBTamilNaduPdfGroupBands(rows, tableStart, headerBandEnd, colCount).forEach((band) => {
      const exists = groupBands.some(
        (b) => Number(b.start) === Number(band.start) && Number(b.end) === Number(band.end)
      );
      if (!exists) groupBands.push(band);
    });
  }
  if (headerModel.formXVIIITamilNadu === true) {
    detectFormXVIIITamilNaduPdfGroupBands(rows, tableStart, headerBandEnd, colCount).forEach(
      (band) => {
        const exists = groupBands.some(
          (b) =>
            Number(b.start) === Number(band.start) && Number(b.end) === Number(band.end)
        );
        if (!exists) groupBands.push(band);
      }
    );
  }
  if (isFormXXIIIAPSE) {
    const xxiiiDeductions = findFormXXIIIAPSEDeductionsBand(rows, tableStart, colCount);
    if (
      xxiiiDeductions &&
      !groupBands.some(
        (band) =>
          band.labelRow === xxiiiDeductions.parentRow &&
          band.start === xxiiiDeductions.start &&
          band.end === xxiiiDeductions.end
      )
    ) {
      const label = String(
        rows[xxiiiDeductions.parentRow]?.[xxiiiDeductions.start] ||
          'Deductions if Any, and Reasons Thereof'
      ).trim();
      groupBands.push({
        labelRow: xxiiiDeductions.parentRow,
        start: xxiiiDeductions.start,
        end: xxiiiDeductions.end,
        label
      });
    }
  }
  // Resolve leaf headers for ALL forms so widths follow column names (Form 11, Form W, …).
  const leafHeaders = resolveLeafHeaderTexts(rows, tableStart, headerBandEnd, colCount);
  const formXVIRemarksColumn = isFormXVIAP
    ? rows
        .slice(Math.max(0, tableStart - 4), Math.min(rows.length, tableStart + 4))
        .reduce((found, row) => {
          if (found >= 0) return found;
          return (row || []).findIndex((cell) => /^remarks?$/i.test(String(cell || '').trim()));
        }, -1)
    : -1;
  const formXVIIHeaderRows = isFormXVIIAP
    ? rows
        .map((row, index) => ({ row, index }))
        .filter(({ row, index }) => {
          if (index > tableStart + 5 || index < Math.max(0, tableStart - 6)) return false;
          const blob = (row || []).join(' ').toLowerCase();
          return (
            /serial\s+no|s\.?\s*no|register\s+of\s+workmen|name\s+of\s+workmen|no\.?\s*of\s*days\s+worked/.test(blob) ||
            /basic\s+wages?|other\s+cash\s+payments?|total\s+deductions?|deductions\s*,?\s*if\s+any/.test(blob)
          );
        })
    : [];
  const formXVIIHeaderTextAt = (column) =>
    formXVIIHeaderRows
      .map(({ row }) => String(row?.[column] || '').replace(/\s+/g, ' ').trim())
      .filter(Boolean)
      .join(' ');
  const formXVIIDeductionStart = isFormXVIIAP
    ? formXVIIHeaderRows.reduce((found, { row }) => {
        const index = (row || []).findIndex((cell) =>
          /deductions\s*,?\s*if\s+any/i.test(String(cell || ''))
        );
        return found >= 0 ? found : index;
      }, -1)
    : -1;
  const isFormXIXRajasthanOvertimeSheet = looksLikeFormXIXRajasthanOvertimePdfContext(
    metaLines,
    rows,
    matrix.name || ''
  );
  const isFormASheet = looksLikeFormAPdfContext(metaLines, rows, matrix.name || '');
  const isFormBRajasthanSheet = looksLikeFormBRajasthanPdfContext(
    metaLines,
    rows,
    matrix.name || '',
    matrix.fileName || ''
  );

  const isFormXVAP = headerModel.formXVAPService === true;
  const isApForm =
    isFormXXAP ||
    isFormXXIAP ||
    isFormXVIIAP ||
    isFormXVIAP ||
    isFormXIIIAP ||
    isFormXIXAP ||
    isFormXVAP ||
    isFormXXIIIAPSE;
  const isApMultiLevelForm = isFormXXAP || isFormXXIAP || isFormXVIIAP || isFormXVIAP;
  const weights = [];
  for (let c = 0; c < colCount; c += 1) {
    const inDayBand = dayBand && c >= dayBand.start && c <= dayBand.end;
    if (inDayBand) {
      weights.push(
        isApMultiLevelForm
          ? 1.6
          : isFormTKASheet
            ? 1.8
            : headerModel.isFormXXVI
              ? 0.42
              : 2.2
      );
      continue;
    }
    if (isFormTKASheet && c === 1) {
      weights.push(8.5);
      continue;
    }
    if (isFormTKASheet && c === 2) {
      weights.push(6);
      continue;
    }
    let maxLen = 4;
    let maxDataLen = 0;
    for (let r = tableStart; r < Math.min(rows.length, tableStart + 80); r += 1) {
      if (isSystemGeneratedDocumentNoteRow(rows[r])) continue;
      if (isUnpaidAccumulationsFootnoteRow(rows[r])) continue;
      if (isForm14RajasthanFootnoteRow(rows[r])) continue;
      if (isFormDRJOuterFootnoteRow(rows[r])) continue;
      const len = String(rows[r][c] || '').length;
      maxLen = Math.max(maxLen, len);
      if (r > headerBandEnd) maxDataLen = Math.max(maxDataLen, len);
    }
    if (headerModel.isFormW) {
      weights.push(formWTamilNaduColumnWeight(leafHeaders[c], maxDataLen || maxLen));
      continue;
    }
    if (headerModel.formBTamilNadu) {
      weights.push(formBTamilNaduPdfColumnWeight(leafHeaders[c]));
      continue;
    }
    if (isFormXXIIIAPSE) {
      weights.push(formXXIIIAPSEColumnWeight(leafHeaders[c]));
      continue;
    }
    if (looksLikeFormXXIIIGJOtPdfContext(metaLines, rows, matrix.name || '', matrix.fileName || '')) {
      const gjW = formXXIIIGJOtPdfColumnWeight(leafHeaders[c], maxDataLen || maxLen);
      if (gjW > 0) {
        weights.push(gjW);
        continue;
      }
    }
    if (headerModel.formLGJGujarat === true) {
      weights.push(formLGJGujaratColumnWeight(leafHeaders[c], maxDataLen || maxLen));
      continue;
    }
    if (headerModel.formPGJGujarat === true) {
      weights.push(formPGJGujaratColumnWeight(leafHeaders[c], maxDataLen || maxLen));
      continue;
    }
    if (headerModel.formQMaharashtra === true) {
      weights.push(formQMaharashtraColumnWeight(leafHeaders[c], maxDataLen || maxLen));
      continue;
    }
    if (headerModel.isFormXXVIIRegister) {
      weights.push(formXXVIITamilNaduColumnWeight(leafHeaders[c], maxDataLen || maxLen));
      continue;
    }
    if (headerModel.isFormXXVI) {
      weights.push(formXXVITamilNaduColumnWeight(leafHeaders[c], maxDataLen || maxLen));
      continue;
    }
    if (
      isApMultiLevelForm &&
      (c === formXVIRemarksColumn ||
        /^remarks?$/i.test(String(leafHeaders[c] || '').replace(/\s+/g, ' ').trim()))
    ) {
      weights.push(isFormXVIAP ? 60 : 9.5);
      continue;
    }
    if (isFormXVIAP) {
      if (c === 0) {
        weights.push(4.5);
        continue;
      }
      if (c === 1 || c === 2) {
        weights.push(13);
        continue;
      }
      if (c === 3) {
        weights.push(3.5);
        continue;
      }
    }
    const weight = statutoryHeaderColumnWeight(leafHeaders[c], maxDataLen || maxLen, colCount);
    const hdrNorm = String(leafHeaders[c] || '')
      .replace(/\s+/g, ' ')
      .trim()
      .toLowerCase();
    // Form A GJ / Form A RJ: keep UAN / Aadhaar readable — same font size needs a wider column.
    if (
      (looksLikeFormAGJGujaratPdfContext(metaLines, rows, matrix.name || '') || isFormASheet) &&
      (/^uan\b/.test(hdrNorm) || /\buan\b/.test(hdrNorm) || /aadha?ar/.test(hdrNorm))
    ) {
      weights.push(Math.max(weight, 11));
      continue;
    }
    weights.push(
      (isFormXIXRajasthanOvertimeSheet || isFormASheet || isFormBRajasthanSheet || isFormXXIIIAPSE) &&
        isPdfSerialNumberHeader(leafHeaders[c])
        ? 5.5
        : weight
    );
  }
  const weightSum = weights.reduce((a, b) => a + b, 0) || 1;
  const rawWidths = weights.map((w) => (w / weightSum) * usableWidth);
  const colWidths = enforcePdfColumnMinWidths(rawWidths, leafHeaders, usableWidth, {
    compactSerial:
      isFormXVIAP ||
      isFormXIXRajasthanOvertimeSheet ||
      isFormASheet ||
      isFormBRajasthanSheet ||
      isFormXXIIIAPSE ||
      isFormXXIIIGJOt,
    tightSerial: isFormXVIAP,
    compactSex: isFormXVIAP || isFormXXIIIGJOt,
    formXXVIWordSafe: headerModel.isFormXXVI === true
  });
  if (headerModel.isFormXXVI && dayBand) {
    const redistributed = redistributeFormXXVIPdfColumnWidths(
      colWidths,
      leafHeaders,
      dayBand,
      usableWidth
    );
    for (let i = 0; i < colWidths.length; i += 1) colWidths[i] = redistributed[i] ?? colWidths[i];
  }
  const colXs = [marginX];
  for (let i = 0; i < colWidths.length; i += 1) colXs.push(colXs[i] + colWidths[i]);
  if (isFormXIVKALayout && colCount === 2) {
    colWidths[0] = usableWidth * 0.44;
    colWidths[1] = usableWidth - colWidths[0];
    colXs.length = 0;
    colXs.push(marginX, marginX + colWidths[0], marginX + usableWidth);
  }
  if (isFormQKALayout && colCount === 2) {
    colWidths[0] = usableWidth * 0.52;
    colWidths[1] = usableWidth - colWidths[0];
    colXs.length = 0;
    colXs.push(marginX, marginX + colWidths[0], marginX + usableWidth);
  }

  const fontSize = isApForm
    ? 9
    : looksLikeFormDRajasthanPdfContext(metaLines, rows, matrix.name || '')
      ? 8
    : isFormBRajasthanSheet
      ? 8
    : looksLikeFormBGJGujaratPdfContext(metaLines, rows, matrix.name || '')
      ? colCount > 22
        ? 7.5
        : 8.5
    : looksLikeFormAGJGujaratPdfContext(metaLines, rows, matrix.name || '')
      ? colCount > 22
        ? 8
        : 8.5
    : isFormASheet
      ? 8
    : headerModel.isFormW
    ? colCount > 28
      ? 5.2
      : 5.8
    : headerModel.isFormXXVIIRegister
      ? 8
      : headerModel.isFormXXVI
        ? 8
      : isFormTKASheet
        ? 7.5
      : isFormQKALayout
        ? 9
      : headerModel.formPGJGujarat === true
        ? 8
      : headerModel.formQMaharashtra === true
        ? FORM_Q_MH_PDF_TABLE_FONT_SIZE
      : headerModel.formXVIIITamilNadu === true
        ? FORM_XVIII_TN_PDF_TABLE_FONT_SIZE
      : colCount > 40
        ? 4.5
        : colCount > 28
          ? 5
          : colCount > 18
            ? 5.5
            : colCount > 14
              ? 6
              : colCount > 10
                ? 7
                : 8;
  let y = startY;
  let pendingSystemNote = false;
  let pendingUnpaidFootnote = '';
  let pendingWhereverApplicableFootnote = '';
  const pendingForm14RjFootnotes = [];
  const pendingFormDRJFootnotes = [];
  const pendingFormDGJFootnotes = [];
  const pendingFormAGJFootnotes = [];

  const looksLikeFormCLwfSheet = (() => {
    const blob = [...metaLines, ...(rows || []).slice(0, Math.min(rows.length, tableStart + 4)).flat(), matrix.name || '']
      .join(' ')
      .toLowerCase();
    return (
      (/form\s*-?\s*c\b/.test(blob) || /lwf\b/.test(blob)) &&
      (/unpaid\s+accumulations|details\s+of\s+fines|quarter\s+ending/.test(blob) ||
        isFormCLwfColHeaderBlob(blob))
    );
  })();
  const looksLikeFormCRajasthanSheet = looksLikeFormCRajasthanPdfContext(
    metaLines,
    rows,
    matrix.name || ''
  );
  const looksLikeFormCGJSheet = looksLikeFormCGJLoanRecoveriesPdfContext(
    metaLines,
    rows,
    matrix.name || ''
  );
  const looksLikeForm14RjSheet = looksLikeForm14RajasthanPdfContext(
    metaLines,
    rows,
    matrix.name || ''
  );
  const looksLikeFormDRJSheet = looksLikeFormDRajasthanPdfContext(
    metaLines,
    rows,
    matrix.name || ''
  );
  const looksLikeFormDGJSheet = looksLikeFormDGJMusterRollPdfContext(
    metaLines,
    rows,
    matrix.name || ''
  );
  const looksLikeFormAGJSheet = looksLikeFormAGJGujaratPdfContext(
    metaLines,
    rows,
    matrix.name || ''
  );
  const looksLikeFormBGJSheet = looksLikeFormBGJGujaratPdfContext(
    metaLines,
    rows,
    matrix.name || ''
  );

  const pushForm14RjFootnote = (raw) => {
    const text = normalizeForm14RajasthanPdfFootnoteText(raw);
    if (!text) return;
    if (pendingForm14RjFootnotes.some((t) => t.toLowerCase() === text.toLowerCase())) return;
    // Keep overtime footnote before the day-entries note.
    if (isForm14RajasthanDayEntriesNoteText(text)) {
      pendingForm14RjFootnotes.push(text);
      return;
    }
    const dayIdx = pendingForm14RjFootnotes.findIndex((t) =>
      isForm14RajasthanDayEntriesNoteText(t)
    );
    if (dayIdx >= 0) pendingForm14RjFootnotes.splice(dayIdx, 0, text);
    else pendingForm14RjFootnotes.push(text);
  };

  const pushFormDRJFootnote = (raw) => {
    const text =
      extractFormDRJOuterFootnoteText(Array.isArray(raw) ? raw : [String(raw || '').trim()]) ||
      normalizeFormDRJOuterFootnoteText(raw);
    if (!text || !isFormDRJOuterFootnoteText(text)) return;
    const key = normalizeFormDRJOuterFootnoteText(text).toLowerCase();
    if (pendingFormDRJFootnotes.some((t) => normalizeFormDRJOuterFootnoteText(t).toLowerCase() === key)) {
      return;
    }
    if (isFormDRJRelayMinesFootnoteText(text)) {
      pendingFormDRJFootnotes.push(FORM_D_RJ_RELAY_MINES_NOTE);
      return;
    }
    if (isFormDRJAbsenceCodesFootnoteText(text)) {
      pendingFormDRJFootnotes.push(FORM_D_RJ_ABSENCE_CODES_NOTE);
      return;
    }
    if (isFormDRJEFormMaintenanceFootnoteText(text)) {
      pendingFormDRJFootnotes.push(FORM_D_RJ_E_FORM_NOTE);
      return;
    }
    pendingFormDRJFootnotes.push(text);
  };

  const pushFormDGJFootnote = (raw) => {
    const text = extractFormDGJGujaratOuterFootnoteText(
      Array.isArray(raw) ? raw : [String(raw || '').trim()]
    ) || String(raw || '').replace(/\s+/g, ' ').trim();
    if (!text || !isFormDGJGujaratOuterFootnoteText(text)) return;
    const key = normalizeFormDGJFootnoteCompare(text);
    if (pendingFormDGJFootnotes.some((t) => normalizeFormDGJFootnoteCompare(t) === key)) return;
    pendingFormDGJFootnotes.push(
      /^\*\*/.test(text.trim())
        ? FORM_DGJ_GJ_ELECTRONIC_FORMAT_NOTE_2
        : /^\*/.test(text.trim()) && /not\s+necessary/i.test(text)
          ? FORM_DGJ_GJ_ELECTRONIC_FORMAT_NOTE_1
          : /governor/i.test(text)
            ? FORM_DGJ_GJ_GOVERNOR_ORDER_NOTE
            : text
    );
  };

  const pushFormAGJFootnote = (raw) => {
    const text =
      extractFormAGJGujaratOuterFootnoteText(
        Array.isArray(raw) ? raw : [String(raw || '').trim()]
      ) || String(raw || '').replace(/\s+/g, ' ').trim();
    if (!text || !isFormAGJGujaratOuterFootnoteText(text)) return;
    const key = normalizeFormAGJFootnoteCompare(text);
    if (pendingFormAGJFootnotes.some((t) => normalizeFormAGJFootnoteCompare(t) === key)) return;
    pendingFormAGJFootnotes.push(text);
  };

  const paintMetaLine = (text, { bold = false, size = 9, align = 'left' } = {}) => {
    const line = String(text || '').replace(/\s+/g, ' ').trim();
    if (!line) return;
    const effectiveSize = isApForm ? 9 : size;
    doc.setFont('helvetica', bold ? 'bold' : 'normal');
    doc.setFontSize(effectiveSize);
    doc.setTextColor(0, 0, 0);
    const wrapped = doc.splitTextToSize(line, usableWidth);
    const blockH = wrapped.length * (effectiveSize + 2) + 3;
    if (y + blockH > pageHeight - marginBottom) {
      doc.addPage();
      y = marginTop;
    }
    if (align === 'center') {
      doc.text(wrapped, pageWidth / 2, y + effectiveSize, { align: 'center' });
    } else if (align === 'right') {
      doc.text(wrapped, pageWidth - marginX, y + effectiveSize, { align: 'right' });
    } else {
      doc.text(wrapped, marginX, y + effectiveSize);
    }
    y += blockH;
  };

  const isDayBandLabelHeaderRow = (row, rowIndex) => {
    if (!dayBand || rowIndex > headerBandEnd) return false;
    if (dayBand.labelRow >= 0) return rowIndex === dayBand.labelRow;
    const blob = (row || []).map((c) => String(c || '')).join(' ').toLowerCase();
    // Fallback only when the leaf label row carries identity + wage headers together.
    return (
      /daily\s+hours\s+of\s+work/.test(blob) ||
      (/rate\s+of\s+wages?/.test(blob) && /name of the work(?:er|man)/.test(blob))
    );
  };

  const groupBandAt = (rowIndex, col) => {
    if (!groupBands.length || rowIndex > headerBandEnd) return null;
    return (
      groupBands.find(
        (b) => b.labelRow === rowIndex && col >= b.start && col <= b.end
      ) || null
    );
  };

  const measureRowHeight = (row, rowIndex) => {
    const mergeDayLabel = isDayBandLabelHeaderRow(row, rowIndex);
    const nilSpan = resolveNilOfTheMonthPdfSpan(row, colCount, rowIndex, headerBandEnd);
    const tnNetSpan = isFormXIXTamilNadu
      ? resolveFormXIXTamilNaduNetAmountPdfSpan(row, colCount)
      : null;
    const isHeaderRow = rowIndex <= headerBandEnd;
    let maxLines = 1;

    if (nilSpan) {
      const bandW = colXs[nilSpan.end + 1] - colXs[nilSpan.start] - 3;
      const wrapped = doc.splitTextToSize(nilSpan.text, Math.max(bandW, 6));
      maxLines = Math.max(1, Math.min(wrapped.length, 3));
      return Math.max(fontSize + 5, maxLines * (fontSize + 1.5) + 4);
    }

    if (tnNetSpan) {
      const labelW = colXs[tnNetSpan.labelEnd + 1] - colXs[tnNetSpan.labelStart] - 3;
      const labelWrapped = doc.splitTextToSize(tnNetSpan.label || ' ', Math.max(labelW, 6));
      maxLines = Math.max(1, Math.min(labelWrapped.length, 4));
      if (tnNetSpan.amountCol >= 0 && tnNetSpan.amount) {
        const amountW = Math.max(colWidths[tnNetSpan.amountCol] - 3, 6);
        const amountWrapped = doc.splitTextToSize(tnNetSpan.amount, amountW);
        maxLines = Math.max(maxLines, Math.min(amountWrapped.length, 3));
      }
      return Math.max(fontSize + 5, maxLines * (fontSize + 1.5) + 4);
    }

    for (let c = 0; c < colCount; c += 1) {
      if (mergeDayLabel && dayBand && c > dayBand.start && c <= dayBand.end) continue;
      const groupBand = groupBandAt(rowIndex, c);
      if (groupBand && c > groupBand.start && c <= groupBand.end) continue;
      const width =
        mergeDayLabel && dayBand && c === dayBand.start
          ? colXs[dayBand.end + 1] - colXs[dayBand.start] - 3
          : groupBand && c === groupBand.start
            ? colXs[groupBand.end + 1] - colXs[groupBand.start] - 3
            : Math.max(colWidths[c] - 3, 6);
      const text =
        mergeDayLabel && dayBand && c === dayBand.start
          ? dayBand.label
          : groupBand && c === groupBand.start
            ? groupBand.label
            : String(row[c] ?? '') || ' ';
      // Data rows: never wrap pure numbers (they shrink-to-fit when painted),
      // except Form A UAN / Aadhaar / long IDs which keep the same font and wrap.
      if (!isHeaderRow && isPurePdfNumericText(text)) {
        const hdr = String(leafHeaders[c] || '')
          .replace(/\s+/g, ' ')
          .trim()
          .toLowerCase();
        const keepFormAIdWrap =
          (looksLikeFormAGJSheet || isFormASheet) &&
          (/^uan\b/.test(hdr) ||
            /\buan\b/.test(hdr) ||
            /aadha?ar/.test(hdr) ||
            /^esic/.test(hdr) ||
            /^lwf\b/.test(hdr) ||
            /^\d{10,}$/.test(String(text || '').replace(/\D/g, '')));
        if (!keepFormAIdWrap) {
          maxLines = Math.max(maxLines, 1);
          continue;
        }
      }
      const wrapped =
        headerModel.isFormXXVI && isHeaderRow
          ? wrapPdfTextPreferWholeWords(doc, text, Math.max(width, 6), 10)
          : doc.splitTextToSize(text, Math.max(width, 6));
      const lineCap =
        groupBand && isFormVIFestivalGroupLabel(groupBand.label)
          ? 10
          : groupBand && isFormBTamilNaduAmountsDeductedGroupLabel(groupBand.label)
            ? 8
          : isFormQKALayout
            ? 12
          : isHeaderRow
            ? headerModel.isFormW
              ? 6
              : headerModel.isFormXXVI
                ? 10
              : String(text || '').length >= 40
                ? 14
                : 8
            : 8;
      maxLines = Math.max(maxLines, Math.min(wrapped.length, lineCap));
    }
    const rowHeight = Math.max(fontSize + 5, maxLines * (fontSize + 1.5) + 4);
    if (headerModel?.formLGJGujarat && isHeaderRow) return Math.max(22, rowHeight);
    if (headerModel?.formPGJGujarat && isHeaderRow) return Math.max(24, rowHeight);
    if (headerModel?.formQMaharashtra && isHeaderRow) return Math.max(56, rowHeight);
    if (headerModel?.formBTamilNadu && isHeaderRow) return Math.max(28, rowHeight);
    // Form XXVI: tall header boxes so whole-word stacks (Permanent / Home / Address) fit.
    if (headerModel?.isFormXXVI && isHeaderRow) {
      return Math.max(42, maxLines * (fontSize + 2.2) + 10);
    }
    // Form X_RJ: tall header boxes so wrapped labels like
    // "Sl. No. of the register of workman employed" are not clipped.
    if (headerModel?.formXRajasthanEmployment && isHeaderRow) {
      return Math.max(46, maxLines * (fontSize + 2.4) + 12);
    }
    return isFormXVIAP ? Math.max(34, rowHeight * 1.35) : rowHeight;
  };

  const paintGridRow = (row, rowIndex, rowH) => {
    const bold = isFormXIXTamilNadu
      ? false
      : isFormQKALayout
        ? false
        : rowIndex <= headerBandEnd || isLikelyHeaderBandRow(row, rowIndex);
    const mergeDayLabel = isDayBandLabelHeaderRow(row, rowIndex);
    const nilSpan = resolveNilOfTheMonthPdfSpan(row, colCount, rowIndex, headerBandEnd);
    const tnNetSpan = isFormXIXTamilNadu
      ? resolveFormXIXTamilNaduNetAmountPdfSpan(row, colCount)
      : null;
    const isHeaderRow = !isFormQKALayout && !isFormXIXTamilNadu && rowIndex <= headerBandEnd;
    const isFormXIIINumberRow =
      isFormXIIIAP &&
      isHeaderRow &&
      (row || []).filter((cell) => /^\(?\s*\d{1,2}\s*\)?$/.test(String(cell || '').trim())).length >= 3;
    doc.setDrawColor(0, 0, 0);
    doc.setLineWidth(0.4);
    doc.setFont('helvetica', bold ? 'bold' : 'normal');
    doc.setFontSize(fontSize);
    doc.setTextColor(0, 0, 0);

    if (nilSpan) {
      const bandW = colXs[nilSpan.end + 1] - colXs[nilSpan.start];
      doc.rect(colXs[nilSpan.start], y, bandW, rowH, 'S');
      const lines = doc
        .splitTextToSize(nilSpan.text, Math.max(bandW - 4, 8))
        .slice(0, 3);
      const textH = lines.length * (fontSize + 1);
      doc.text(lines, colXs[nilSpan.start] + bandW / 2, y + (rowH - textH) / 2 + fontSize, {
        align: 'center'
      });
      y += rowH;
      return;
    }

    if (tnNetSpan) {
      const labelW = colXs[tnNetSpan.labelEnd + 1] - colXs[tnNetSpan.labelStart];
      doc.rect(colXs[tnNetSpan.labelStart], y, labelW, rowH, 'S');
      const labelLines = doc
        .splitTextToSize(tnNetSpan.label || ' ', Math.max(labelW - 3, 6))
        .slice(0, 4);
      doc.text(labelLines, colXs[tnNetSpan.labelStart] + 1.5, y + fontSize + 1);
      if (tnNetSpan.amountCol >= 0) {
        doc.rect(colXs[tnNetSpan.amountCol], y, colWidths[tnNetSpan.amountCol], rowH, 'S');
        if (tnNetSpan.amount) {
          const amountLines = doc
            .splitTextToSize(tnNetSpan.amount, Math.max(colWidths[tnNetSpan.amountCol] - 3, 6))
            .slice(0, 3);
          doc.text(amountLines, colXs[tnNetSpan.amountCol] + 1.5, y + fontSize + 1);
        }
      }
      y += rowH;
      return;
    }

    for (let c = 0; c < colCount; c += 1) {
      if (mergeDayLabel && dayBand && c > dayBand.start && c <= dayBand.end) continue;
      const groupBand = groupBandAt(rowIndex, c);
      if (groupBand && c > groupBand.start && c <= groupBand.end) continue;

      if (mergeDayLabel && dayBand && c === dayBand.start) {
        const bandW = colXs[dayBand.end + 1] - colXs[dayBand.start];
        doc.rect(colXs[c], y, bandW, rowH, 'S');
        const lines = doc
          .splitTextToSize(dayBand.label, Math.max(bandW - 4, 8))
          .slice(0, 3);
        const textH = lines.length * (fontSize + 1);
        doc.text(lines, colXs[c] + bandW / 2, y + (rowH - textH) / 2 + fontSize, {
          align: 'center'
        });
        continue;
      }

      if (groupBand && c === groupBand.start) {
        const bandW = colXs[groupBand.end + 1] - colXs[groupBand.start];
        doc.rect(colXs[c], y, bandW, rowH, 'S');
        const groupLineCap = isFormVIFestivalGroupLabel(groupBand.label)
          ? 10
          : isFormBTamilNaduAmountsDeductedGroupLabel(groupBand.label)
            ? 8
            : 3;
        const lines = doc
          .splitTextToSize(groupBand.label, Math.max(bandW - 4, 8))
          .slice(0, groupLineCap);
        const textH = lines.length * (fontSize + 1);
        doc.text(lines, colXs[c] + bandW / 2, y + (rowH - textH) / 2 + fontSize, {
          align: 'center'
        });
        continue;
      }

      doc.rect(colXs[c], y, colWidths[c], rowH, 'S');
      let raw = String(row[c] ?? '');
      if (isFormXXIIIGJOt && !isHeaderRow) {
        raw = String(
          sanitizeFormXXIIIGJOtMatrixCell(raw, leafHeaders[c] || '', c) ?? ''
        );
      }
      if (!raw.trim()) continue;
      // Hide duplicate day-band label leftovers in non-label header rows
      if (
        dayBand &&
        rowIndex <= headerBandEnd &&
        c >= dayBand.start &&
        c <= dayBand.end &&
        (/daily\s+hours/i.test(raw) ||
          (isFormTKASheet && /attendance/i.test(raw) && !/^\d{1,2}$/.test(raw.trim())) ||
          (headerModel.formPGJGujarat === true && isFormPGJDateOfMonthGroupLabel(raw)) ||
          (headerModel.formQMaharashtra === true && isFormQMaharashtraParentGroupLabel(raw)))
      ) {
        continue;
      }
      // Form L GJ: never paint "Weekly holiday day" inside the Date of the Month band.
      if (
        headerModel.formLGJGujarat === true &&
        rowIndex <= headerBandEnd &&
        isFormLGJWeeklyHolidayHeaderText(raw)
      ) {
        const dateBand = formLGJGroupBands.find(
          (b) => b && isFormLGJDateOfMonthGroupLabel(b.label) && b.end > b.start
        );
        if (dateBand && c >= dateBand.start && c <= dateBand.end) {
          continue;
        }
      }
      // Hide leftover group labels that leaked into leaf header cells
      if (
        groupBands.length &&
        rowIndex <= headerBandEnd &&
        isStatutoryGroupHeaderLabel(raw) &&
        !groupBandAt(rowIndex, c)
      ) {
        continue;
      }
      // Form Q MH: parent banners only paint via groupBand merge — never as leaf text.
      if (
        headerModel.formQMaharashtra === true &&
        rowIndex <= headerBandEnd &&
        isFormQMaharashtraParentGroupLabel(raw) &&
        !groupBandAt(rowIndex, c)
      ) {
        continue;
      }
      // Form Q MH: never paint stray Advances/Other under Total Deduction / Net Payable.
      if (headerModel.formQMaharashtra === true && rowIndex <= headerBandEnd) {
        const stackParts = [];
        for (let hr = Math.max(0, tableStart - 2); hr <= headerBandEnd; hr += 1) {
          const ht = String(rows[hr]?.[c] ?? '')
            .replace(/\s+/g, ' ')
            .trim();
          if (ht) stackParts.push(ht);
        }
        const colStack = stackParts.join(' ').toLowerCase();
        if (
          /total\s+deductions?|net\s+payable/.test(colStack) &&
          (/^advances?\b/i.test(raw.trim()) || /other\s+deductions?/i.test(raw)) &&
          !/total\s+deductions?|net\s+payable/i.test(raw)
        ) {
          continue;
        }
      }

      // Form Q MH: heading-box labels (Name of the worker, etc.) — bottom→top (angle 90).
      // Long labels wrap to 2 lines so they fit the header box height.
      if (
        headerModel.formQMaharashtra === true &&
        isHeaderRow &&
        isFormQMaharashtraVerticalHeaderText(raw)
      ) {
        const label = String(raw || '')
          .replace(/\s+/g, ' ')
          .trim();
        // Span empty header rows below so long labels fit (identity merges).
        let spanH = rowH;
        for (let hr = rowIndex + 1; hr <= headerBandEnd; hr += 1) {
          const below = String(rows[hr]?.[c] ?? '')
            .replace(/\s+/g, ' ')
            .trim();
          if (below) break;
          spanH += measureRowHeight(rows[hr], hr);
        }
        const colW = Math.max(colWidths[c] - 2, 6);
        let lines = splitFormQMaharashtraVerticalHeaderLines(label, 2);
        let size = Math.min(fontSize, 7);
        doc.setFont('helvetica', 'bold');
        doc.setFontSize(size);
        const maxLen = Math.max(spanH - 6, 10);
        const longest = () => Math.max(...lines.map((ln) => doc.getTextWidth(ln)), 0);
        const linesNeedWidth = () =>
          lines.length <= 1 ? 0 : (lines.length - 1) * (size + 1.2) + size;
        while (size > 3.6 && (longest() > maxLen || linesNeedWidth() > colW)) {
          size -= 0.25;
          doc.setFontSize(size);
        }
        const tw = longest();
        const blockW = lines.length > 1 ? (lines.length - 1) * (size + 1.2) : 0;
        const cx = colXs[c] + colWidths[c] / 2;
        const cy = y + spanH / 2;
        // jsPDF angle 90 = counterclockwise → reads bottom→top; multi-line spreads across column.
        doc.text(lines, cx + size * 0.35 - blockW / 2, cy + tw / 2, { angle: 90 });
        doc.setFontSize(fontSize);
        doc.setFont('helvetica', isHeaderRow ? 'bold' : 'normal');
        continue;
      }

      const cellW = Math.max(colWidths[c] - 3, 6);
      const alignRight = isPurePdfNumericText(raw);
      const normalizedLeaf = String(leafHeaders[c] || '')
        .replace(/\s+/g, ' ')
        .trim()
        .toLowerCase();
      const formXVIIHeaderText = `${normalizedLeaf} ${formXVIIHeaderTextAt(c)}`.toLowerCase();
      const formXVIICenterColumn =
        isFormXVIIAP &&
        (c === 0 ||
          /days?\s+worked|serial\s+no/.test(formXVIIHeaderText) ||
          (formXVIIDeductionStart >= 0 &&
            c >= formXVIIDeductionStart &&
            c < formXVIIDeductionStart + 5));

      const formXVIIAmountColumn =
        isFormXVIIAP &&
        /basic\s+wages?|other\s+cash\s+payments?|total\s+deductions?/.test(formXVIIHeaderText);
      const displayRaw = formXVIIAmountColumn && /^-?\d+(?:\.\d+)?$/.test(raw.trim())
        ? Number(raw).toLocaleString('en-IN', { maximumFractionDigits: 0 })
        : raw;

      if (isFormXIXAP && colCount === 2) {
        const apDisplay = c === 1
          ? formatFormXIXAPMoneyValue(row[0], raw)
          : raw;
        const apLines = doc.splitTextToSize(apDisplay, cellW).slice(0, 8);
        if (c === 1 && isPurePdfNumericText(apDisplay)) {
          doc.text(apLines, colXs[c] + colWidths[c] - 1.5, y + fontSize + 1, { align: 'right' });
        } else {
          doc.text(apLines, colXs[c] + 1.5, y + fontSize + 1);
        }
        continue;
      }

      // Tamil Nadu Form XIX: all labels and values left-aligned, never bold (set above).
      if (isFormXIXTamilNadu) {
        const tnLines = doc.splitTextToSize(raw, cellW).slice(0, 8);
        doc.text(tnLines, colXs[c] + 1.5, y + fontSize + 1);
        continue;
      }

      if (isFormXIXKA) {
        const kaLines = doc.splitTextToSize(raw, cellW).slice(0, 6);
        const textH = kaLines.length * (fontSize + 1);
        doc.text(kaLines, colXs[c] + colWidths[c] / 2, y + (rowH - textH) / 2 + fontSize, {
          align: 'center'
        });
        continue;
      }

      if (isFormXIIINumberRow && alignRight) {
        doc.text(raw, colXs[c] + colWidths[c] / 2, y + (rowH + fontSize) / 2 - 1, {
          align: 'center'
        });
        continue;
      }

      if (formXVIIAmountColumn && !isHeaderRow && displayRaw.trim()) {
        let size = fontSize;
        doc.setFontSize(size);
        while (!isApForm && size > 3.2 && doc.getTextWidth(displayRaw) > cellW) {
          size -= 0.4;
          doc.setFontSize(size);
        }
        doc.text(displayRaw, colXs[c] + colWidths[c] - 1.5, y + (rowH + size) / 2 - 1, { align: 'right' });
        doc.setFontSize(fontSize);
        continue;
      }

      if (formXVIICenterColumn && !isHeaderRow && displayRaw.trim()) {
        let size = fontSize;
        doc.setFontSize(size);
        while (!isApForm && size > 3.2 && doc.getTextWidth(displayRaw) > cellW) {
          size -= 0.4;
          doc.setFontSize(size);
        }
        doc.text(displayRaw, colXs[c] + colWidths[c] / 2, y + (rowH + size) / 2 - 1, { align: 'center' });
        doc.setFontSize(fontSize);
        continue;
      }

      if (isFormXXIIIAPSE && !isHeaderRow) {
        const xxiiiHeader = leafHeaders[c] || '';
        const xxiiiDisplay = isFormXXIIIAPSEMoneyHeader(xxiiiHeader)
          ? formatFormXXIIIAPSEMoneyPdfValue(raw)
          : raw;
        if (isFormXXIIIAPSECenterValueHeader(xxiiiHeader) && String(xxiiiDisplay || '').trim()) {
          doc.text(xxiiiDisplay, colXs[c] + colWidths[c] / 2, y + (rowH + fontSize) / 2 - 1, {
            align: 'center'
          });
          continue;
        }
        if (isFormXXIIIAPSEMoneyHeader(xxiiiHeader) && String(xxiiiDisplay || '').trim()) {
          doc.text(xxiiiDisplay, colXs[c] + colWidths[c] - 1.5, y + (rowH + fontSize) / 2 - 1, {
            align: 'right'
          });
          continue;
        }
      }

      if (isFormXXIIIGJOt && !isHeaderRow) {
        const gjHeader = leafHeaders[c] || '';
        if (isFormXXIIIGJOtCenterValueHeader(gjHeader) && String(raw || '').trim()) {
          doc.text(raw, colXs[c] + colWidths[c] / 2, y + (rowH + fontSize) / 2 - 1, {
            align: 'center'
          });
          continue;
        }
      }

      // Form XVI AP: center S.No column (column 0) and date columns (1-31)
      const isFormXVISNoColumn = isFormXVIAP && c === 0;
      const isFormXVIDateColumn = isFormXVIAP && formXVIDateBand && c >= formXVIDateBand.start && c <= formXVIDateBand.end;
     
      if ((isFormXXAP || isFormXXIAP || isFormXIIIAP || isFormXVISNoColumn) && c === 0) {
        doc.text(raw, colXs[c] + colWidths[c] / 2, y + (rowH + fontSize) / 2 - 1, {
          align: 'center'
        });
        continue;
      }

      // Center the date columns in Form XVI AP
      if (isFormXVIDateColumn && !isHeaderRow && raw.trim()) {
        doc.text(raw, colXs[c] + colWidths[c] / 2, y + (rowH + fontSize) / 2 - 1, {
          align: 'center'
        });
        continue;
      }

      // Amounts / counts: single line, shrink font instead of mid-digit wrap.
      // Form A GJ / Form A RJ UAN / Aadhaar / long IDs: keep same Helvetica + fontSize as other cells.
      if (!isHeaderRow && alignRight) {
        const hdr = String(leafHeaders[c] || '')
          .replace(/\s+/g, ' ')
          .trim()
          .toLowerCase();
        const keepFormAIdFont =
          (looksLikeFormAGJSheet || isFormASheet) &&
          (/^uan\b/.test(hdr) ||
            /\buan\b/.test(hdr) ||
            /aadha?ar/.test(hdr) ||
            /^esic/.test(hdr) ||
            /^lwf\b/.test(hdr) ||
            /bank\s*a\s*\/?\s*c|account/.test(hdr) ||
            /^\d{10,}$/.test(String(raw || '').replace(/\D/g, '')));
        if (keepFormAIdFont) {
          doc.setFont('helvetica', 'normal');
          doc.setFontSize(fontSize);
          const idLines = doc.splitTextToSize(raw, cellW).slice(0, 3);
          const textH = idLines.length * (fontSize + 1);
          doc.text(idLines, colXs[c] + 1.5, y + (rowH - textH) / 2 + fontSize);
          continue;
        }
        let size = fontSize;
        doc.setFontSize(size);
        while (
          !isApForm &&
          headerModel.formPGJGujarat !== true &&
          size > 3.2 &&
          doc.getTextWidth(raw) > cellW
        ) {
          size -= 0.4;
          doc.setFontSize(size);
        }
        doc.text(raw, colXs[c] + colWidths[c] - 1.5, y + (rowH + size) / 2 - 1, {
          align: 'right'
        });
        doc.setFontSize(fontSize);
        continue;
      }

      const lineCap = isFormQKALayout
        ? 12
        : isHeaderRow && headerModel.isFormW
          ? 6
          : isHeaderRow && headerModel.isFormXXVI
            ? 10
          : isHeaderRow && raw.length >= 40
            ? 14
            : 8;
      const lines =
        headerModel.isFormXXVI && isHeaderRow
          ? wrapPdfTextPreferWholeWords(doc, raw, cellW, lineCap)
          : doc.splitTextToSize(raw, cellW).slice(0, lineCap);
      if (alignRight && !isFormQKALayout) {
        doc.text(lines, colXs[c] + colWidths[c] - 1.5, y + fontSize + 1, { align: 'right' });
      } else if (!isFormQKALayout && (isHeaderRow || isPureNilPdfText(raw))) {
        // Header labels and Nil/NIL values — center like the Excel register model.
        const textH = lines.length * (fontSize + 1);
        doc.text(lines, colXs[c] + colWidths[c] / 2, y + (rowH - textH) / 2 + fontSize, {
          align: 'center'
        });
      } else {
        doc.text(lines, colXs[c] + 1.5, y + fontSize + 1);
      }
    }
    y += rowH;
  };

  const repeatColumnHeaders = () => {
    for (let hr = tableStart; hr <= headerBandEnd && hr < rows.length; hr += 1) {
      if (!rowHasContent(rows[hr]) && hr < headerBandEnd) continue;
      const rowH = measureRowHeight(rows[hr], hr);
      if (y + rowH > pageHeight - marginBottom) {
        doc.addPage();
        y = marginTop;
      }
      paintGridRow(rows[hr], hr, rowH);
    }
  };

  const paintFormXXAPTableHeader = () => {
    if (!isFormXXAP || tableStart + 1 >= rows.length) return;
    const topRow = rows[tableStart] || [];
    const childRow = rows[tableStart + 1] || [];
    const topH = 46;
    const childH = 32;
    const totalH = topH + childH;
    if (y + totalH > pageHeight - marginBottom) {
      doc.addPage();
      y = marginTop;
    }
    doc.setDrawColor(0, 0, 0);
    doc.setLineWidth(0.4);
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(9);
    for (let c = 0; c < colCount; c += 1) {
      const isRecoveryChild = c >= 9 && c <= 11;
      if (!isRecoveryChild) {
        doc.rect(colXs[c], y, colWidths[c], totalH, 'S');
        const text = String(topRow[c] || '').trim();
        if (text) {
          const lines = /^remarks?$/i.test(text)
            ? [text]
            : doc.splitTextToSize(text, Math.max(colWidths[c] - 4, 8)).slice(0, 5);
          const headerFontSize = isFormXVIAP ? 9 : 11;
          const textH = lines.length * (headerFontSize + 1);
          doc.text(lines, colXs[c] + colWidths[c] / 2, y + (totalH - textH) / 2 + headerFontSize, {
            align: 'center'
          });
        }
        continue;
      }
      if (c === 9) {
        const recoveryW = colXs[12] - colXs[9];
        doc.rect(colXs[9], y, recoveryW, topH, 'S');
        const parent = String(topRow[9] || 'Date of recovery').trim();
        const lines = doc.splitTextToSize(parent, Math.max(recoveryW - 4, 8)).slice(0, 3);
        const textH = lines.length * (fontSize + 1);
        doc.text(lines, colXs[9] + recoveryW / 2, y + (topH - textH) / 2 + fontSize, {
          align: 'center'
        });
      }
      const childText = String(childRow[c] || '').trim();
      doc.rect(colXs[c], y + topH, colWidths[c], childH, 'S');
      if (childText) {
        const lines = doc.splitTextToSize(childText, Math.max(colWidths[c] - 4, 8)).slice(0, 3);
        const textH = lines.length * (fontSize + 1);
        doc.text(lines, colXs[c] + colWidths[c] / 2, y + topH + (childH - textH) / 2 + fontSize, {
          align: 'center'
        });
      }
    }
    y += totalH;
  };

  const paintFormXVIIAPTableHeader = () => {
    if (!isFormXVIIAP) return;
    let parentRowIndex = -1;
    const searchStart = Math.max(0, tableStart - 8);
    const searchEnd = Math.min(rows.length - 1, tableStart + 5);
    for (let r = searchStart; r <= searchEnd; r += 1) {
      if ((rows[r] || []).some((cell) => /deductions\s*,?\s*if\s+any/i.test(String(cell || '')))) {
        parentRowIndex = r;
        break;
      }
    }
    if (parentRowIndex < 0 || parentRowIndex + 1 >= rows.length) return;
    const topRow = rows[parentRowIndex] || [];
    const childRow = rows[parentRowIndex + 1] || [];
    const topH = 52;
    const childH = 30;
    const totalH = topH + childH;
    const headerFontSize = 8;
    const lineHeightFactor = 1.12;
    const deductionStart = topRow.findIndex((cell) =>
      /deductions\s*,?\s*if\s+any/i.test(String(cell || ''))
    );
    if (deductionStart < 0) return;
    const ranges = [{ start: deductionStart, end: Math.min(colCount - 1, deductionStart + 4) }];
    const wageStart = topRow.findIndex((cell) => /wage\s+period/i.test(String(cell || '')));
    if (wageStart >= 0) ranges.unshift({ start: wageStart, end: Math.min(colCount - 1, wageStart + 1) });
    const mergedColumn = (col) => ranges.find((range) => col >= range.start && col <= range.end);
    if (y + totalH > pageHeight - marginBottom) {
      doc.addPage();
      y = marginTop;
    }
    const previousLineHeight =
      typeof doc.getLineHeightFactor === 'function' ? doc.getLineHeightFactor() : 1.15;
    doc.setDrawColor(0, 0, 0);
    doc.setLineWidth(0.4);
    doc.setFont('helvetica', 'bold');
    doc.setLineHeightFactor(lineHeightFactor);

    const paintFormXVIIHeaderLabel = (text, boxX, boxY, boxW, boxH, vAlign = 'top') => {
      const raw = String(text || '').trim();
      if (!raw) return;
      let size = headerFontSize;
      doc.setFontSize(size);
      const maxW = Math.max(boxW - 4.4, 6);
      let lines = doc.splitTextToSize(raw, maxW);
      const maxFit = (s) => Math.max(1, Math.floor((boxH - 13) / Math.max(s * lineHeightFactor, 1)));
      while (size > 6.2 && lines.length > maxFit(size)) {
        size -= 0.3;
        doc.setFontSize(size);
        lines = doc.splitTextToSize(raw, maxW);
      }
      lines = lines.slice(0, maxFit(size));
      doc.setFontSize(size);
      const { firstBaseline } = formXVIIAPHeaderTextMetrics({
        lineCount: lines.length,
        fontSize: size,
        cellHeight: boxH,
        lineHeightFactor,
        vAlign
      });
      doc.text(lines, boxX + boxW / 2, boxY + firstBaseline, { align: 'center' });
    };

    for (let c = 0; c < colCount; c += 1) {
      const range = mergedColumn(c);
      if (range && c > range.start) continue;
      if (range) {
        const width = colXs[range.end + 1] - colXs[range.start];
        doc.rect(colXs[range.start], y, width, topH, 'S');
        paintFormXVIIHeaderLabel(topRow[range.start], colXs[range.start], y, width, topH, 'middle');
        for (let child = range.start; child <= range.end; child += 1) {
          doc.rect(colXs[child], y + topH, colWidths[child], childH, 'S');
          paintFormXVIIHeaderLabel(
            childRow[child],
            colXs[child],
            y + topH,
            colWidths[child],
            childH,
            'middle'
          );
        }
        continue;
      }
      doc.rect(colXs[c], y, colWidths[c], totalH, 'S');
      paintFormXVIIHeaderLabel(topRow[c], colXs[c], y, colWidths[c], totalH, 'top');
    }
    doc.setLineHeightFactor(previousLineHeight);
    doc.setFontSize(fontSize);
    y += totalH;
  };

  const paintFormXVIAPTableHeader = () => {
    if (!isFormXVIAP || !formXVIDateBand) return;
   
    // Find the header row that contains S.No, Name of Employee, etc.
    let headerRowIndex = -1;
    for (let r = Math.max(0, tableStart - 4); r <= Math.min(rows.length - 1, tableStart + 2); r += 1) {
      const blob = (rows[r] || []).join(' ').toLowerCase();
      if (/s\.?\s*no|name\s+of\s+the\s+employee/i.test(blob)) {
        headerRowIndex = r;
        break;
      }
    }
   
    if (headerRowIndex < 0 || headerRowIndex + 1 >= rows.length) return;
   
    const topRow = rows[headerRowIndex] || [];
    // The dateRow is the one that actually contains the date numbers 1-31
    // This is the row AFTER labelRow (which is found as r-1 in formXVIDateBand detection)
    const dateRow = rows[formXVIDateBand.labelRow + 1] || rows[headerRowIndex + 1] || [];
    const topH = 48;
    const dateH = 34;
    const totalH = topH + dateH;
   
    if (y + totalH > pageHeight - marginBottom) {
      doc.addPage();
      y = marginTop;
    }
   
    doc.setDrawColor(0, 0, 0);
    doc.setLineWidth(0.4);
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(9);
   
    // Iterate through columns
    for (let c = 0; c < colCount; c += 1) {
      const inDateBand = c >= formXVIDateBand.start && c <= formXVIDateBand.end;
     
      // Skip if we've already handled this column as part of a merged date band
      if (inDateBand && c > formXVIDateBand.start) continue;
     
      // Dates band: merge all date columns under "Dates" parent
      if (inDateBand && c === formXVIDateBand.start) {
        const dateWidth = colXs[formXVIDateBand.end + 1] - colXs[formXVIDateBand.start];
       
        // Top row: "Dates" header spanning all date columns
        doc.rect(colXs[formXVIDateBand.start], y, dateWidth, topH, 'S');
        doc.setFont('helvetica', 'bold');
        doc.setFontSize(fontSize);
        doc.text('Dates', colXs[formXVIDateBand.start] + dateWidth / 2, y + topH / 2 + fontSize / 2, { align: 'center' });
       
        // Bottom row: individual date numbers (1-31)
        for (let dateCol = formXVIDateBand.start; dateCol <= formXVIDateBand.end; dateCol += 1) {
          doc.rect(colXs[dateCol], y + topH, colWidths[dateCol], dateH, 'S');
          const dateNum = String(dateRow[dateCol] || '').trim();
          if (dateNum && /^\d{1,2}$/.test(dateNum)) {
            doc.setFont('helvetica', 'normal');
            doc.setFontSize(fontSize);
            doc.text(dateNum, colXs[dateCol] + colWidths[dateCol] / 2, y + topH + dateH / 2 + fontSize / 2, { align: 'center' });
          }
        }
        continue;
      }
     
      // Non-date columns: full height (top + date height combined)
      doc.rect(colXs[c], y, colWidths[c], totalH, 'S');
      const text = String(topRow[c] || '').trim();
      if (text) {
        doc.setFont('helvetica', 'bold');
        let headerFontSize = 9;
        doc.setFontSize(headerFontSize);
        const isRemarksHeader = /^remarks?$/i.test(text);
        const availableWidth = Math.max(colWidths[c] - 6, 8);
        while (isRemarksHeader && headerFontSize > 6 && doc.getTextWidth(text) > availableWidth) {
          headerFontSize -= 0.5;
          doc.setFontSize(headerFontSize);
        }
        const lines = isRemarksHeader
          ? [text]
          : doc.splitTextToSize(text, availableWidth).slice(0, 4);
        const textH = lines.length * (headerFontSize + 1);
        doc.text(lines, colXs[c] + colWidths[c] / 2, y + (totalH - textH) / 2 + headerFontSize, { align: 'center' });
      }
    }
   
    y += totalH;
  };

  // All forms: bordered “vertical line” header (form name + fields), matching Excel model.
  if ((headerModel.titles || []).length || (headerModel.fields || []).length) {
    y = paintBorderedStatutoryHeader(
      doc,
      headerModel,
      { pageWidth, pageHeight, marginX, marginTop, marginBottom, usableWidth },
      y
    );
  } else if (
    matrix.name &&
    matrix.name !== 'Sheet1' &&
    !looksLikeExcelDraftFileLabel(matrix.name) &&
    !looksLikeExcelSheetTabName(matrix.name)
  ) {
    y = paintBorderedStatutoryHeader(
      doc,
      { titles: [matrix.name], fields: [], rightFields: [] },
      { pageWidth, pageHeight, marginX, marginTop, marginBottom, usableWidth },
      y
    );
  }

  paintFormXXAPTableHeader();
  paintFormXVIIAPTableHeader();
  paintFormXVIAPTableHeader();

  if (headerModel.hasSystemNote) pendingSystemNote = true;
  metaLines.forEach((line) => {
    if (isSystemGeneratedDocumentNote(line)) pendingSystemNote = true;
    if (isUnpaidAccumulationsFootnoteText(line) && !pendingUnpaidFootnote) {
      pendingUnpaidFootnote = String(line).replace(/\s+/g, ' ').trim();
    }
    if (isFormCWhereverApplicableFootnoteText(line) && !pendingWhereverApplicableFootnote) {
      pendingWhereverApplicableFootnote = FORM_C_GJ_WHEREVER_APPLICABLE_FOOTNOTE;
    }
    if (isForm14RajasthanFootnoteText(line)) pushForm14RjFootnote(line);
    if (isFormDRJOuterFootnoteText(line) && looksLikeFormDRJSheet) {
      pushFormDRJFootnote(line);
    }
    if (isFormDGJGujaratOuterFootnoteText(line) && !looksLikeFormAGJSheet && !looksLikeFormBGJSheet) {
      pushFormDGJFootnote(line);
    }
    if (isFormAGJGujaratOuterFootnoteText(line) && !looksLikeFormBGJSheet) {
      pushFormAGJFootnote(line);
    }
  });

  for (let r = tableStart; r < rows.length; r += 1) {
    const row = rows[r];
    if (!rowHasContent(row)) continue;
    if (isFormXIIIAP && isFormXIIIPdfTitleRow(row)) continue;
    if (isFormXXIIIAPSE && (isFormXXIIIAPSETitleRow(row) || isFormXXIIIAPSEAdminRow(row))) continue;
    if (
      isFormXXIIIGJOt &&
      (isFormXXIIIGJOtTitleRow(row) ||
        isFormXXIIIGJOtAdminRow(row) ||
        isFormXXIIIGJOtPreambleRow(row))
    ) {
      continue;
    }
    if (isFormXXIIIGJOt && r > tableStart + 1 && isFormXXIIIGJOtSpuriousPdfDataRow(row)) {
      continue;
    }
    if (isApForm && isApPdfHeadingRow(row)) continue;
    if ((isFormXXAP || isFormXVIIAP || isFormXVIAP) && r <= headerBandEnd) continue;

    // Never draw the footer inside Sr. No. — paint after the full column band.
    if (isSystemGeneratedDocumentNoteRow(row)) {
      pendingSystemNote = true;
      continue;
    }
    if (isFormTKASheet && isFormTSEKarnatakaPdfFooterOnlyRow(row)) {
      continue;
    }
    if (headerModel?.formBTamilNadu && isFormBTamilNaduPdfSignatoryRow(row)) {
      continue;
    }
    if (
      (headerModel?.formCTamilNaduLwf || looksLikeFormCLwfSheet) &&
      isFormCTamilNaduLwfPdfSignatoryRow(row)
    ) {
      continue;
    }

    // Form C legal footnote — paint below the grid, above the system-generated note.
    if (isUnpaidAccumulationsFootnoteRow(row) || isUnpaidAccumulationsFootnoteText(row?.[0])) {
      const fn = extractUnpaidAccumulationsFootnoteText(row) || String(row?.[0] || '').trim();
      if (fn && !pendingUnpaidFootnote) pendingUnpaidFootnote = fn;
      continue;
    }

    // Form C GJ: keep "*Wherever applicable" outside the table box.
    if (
      isFormCWhereverApplicableFootnoteRow(row) ||
      isFormCWhereverApplicableFootnoteText(row?.[0])
    ) {
      if (!pendingWhereverApplicableFootnote) {
        pendingWhereverApplicableFootnote =
          extractFormCWhereverApplicableFootnoteText(row) ||
          FORM_C_GJ_WHEREVER_APPLICABLE_FOOTNOTE;
      }
      continue;
    }

    // Form C GJ: never print "*Applicable only in case of damage/loss/fine".
    if (
      isFormCDamageLossFineFootnoteRow(row) ||
      isFormCDamageLossFineFootnoteText(row?.[0])
    ) {
      if (looksLikeFormCGJSheet || !looksLikeFormCRajasthanSheet) continue;
      if (!pendingUnpaidFootnote) pendingUnpaidFootnote = FORM_C_RJ_FOOTNOTE;
      continue;
    }

    // Form 14 RJ footnotes — paint below the grid, above the system-generated note.
    if (isForm14RajasthanFootnoteRow(row) || isForm14RajasthanFootnoteText(row?.[0])) {
      const fn = extractForm14RajasthanFootnoteText(row) || String(row?.[0] || '').trim();
      if (fn) pushForm14RjFootnote(fn);
      continue;
    }

    // Form D RJ: keep Mines / absence-code / E-Form notes outside the attendance box.
    if (
      looksLikeFormDRJSheet &&
      (isFormDRJOuterFootnoteRow(row) || isFormDRJOuterFootnoteText(row?.[0]))
    ) {
      pushFormDRJFootnote(row);
      continue;
    }

    // Form D GJ: keep electronic-format / Governor notes outside the muster-roll box.
    if (
      !looksLikeFormAGJSheet &&
      (isFormDGJGujaratOuterFootnoteRow(row) ||
        isFormDGJGujaratOuterFootnoteText(row?.[0]))
    ) {
      pushFormDGJFootnote(row);
      continue;
    }

    // Form A GJ: keep age / skill / electronic / wherever notes outside the employee box.
    if (
      !looksLikeFormBGJSheet &&
      (isFormAGJGujaratOuterFootnoteRow(row) ||
        isFormAGJGujaratOuterFootnoteText(row?.[0]))
    ) {
      pushFormAGJFootnote(row);
      continue;
    }

    // Form B GJ: never paint Rate of Minimum Wages scraps inside the wage table.
    if (
      looksLikeFormBGJSheet &&
      (isFormBGJMinimumWagesMetaText(row?.[0]) ||
        (Array.isArray(row) &&
          row.some((c) => isFormBGJMinimumWagesMetaText(c)) &&
          row.filter((c) => String(c || '').trim()).length <= 6))
    ) {
      continue;
    }

    // Safety: duplicated long meta already shown in the bordered header — skip grid paint
    const filled = row.filter((c) => String(c || '').trim());
    const uniqueFilled = [...new Set(filled)];
    if (
      r <= headerBandEnd + 1 &&
      uniqueFilled.length === 1 &&
      uniqueFilled[0].length > 20 &&
      filled.length >= 3
    ) {
      continue;
    }

    const rowH = measureRowHeight(row, r);
    if (y + rowH > pageHeight - marginBottom) {
      doc.addPage();
      y = marginTop;
      if (r > headerBandEnd) repeatColumnHeaders();
    }
    paintGridRow(row, r, rowH);
  }

  // Form T KA: Date: (left) and Authorised Signatory (right) sit under the table box.
  if (isFormTKASheet) {
    const footer = matrix.formTKAFooter || {
      date: FORM_T_KA_PDF_DATE_LABEL,
      signatory: FORM_T_KA_PDF_SIGNATORY_LABEL,
    };
    const dateText = String(footer.date || FORM_T_KA_PDF_DATE_LABEL).trim() || FORM_T_KA_PDF_DATE_LABEL;
    const signatoryText =
      String(footer.signatory || FORM_T_KA_PDF_SIGNATORY_LABEL).trim() ||
      FORM_T_KA_PDF_SIGNATORY_LABEL;
    const size = 9;
    const h = size + 8;
    y += 10;
    if (y + h > pageHeight - marginBottom) {
      doc.addPage();
      y = marginTop;
    }
    doc.setFont('helvetica', 'normal');
    doc.setFontSize(size);
    doc.setTextColor(0, 0, 0);
    doc.text(dateText, marginX, y + size);
    doc.text(signatoryText, pageWidth - marginX, y + size, { align: 'right' });
    y += h;
  }

  if (isFormXIXAP && Array.isArray(matrix.formXIXAPFooterLines)) {
    matrix.formXIXAPFooterLines.forEach((line) => {
      y += 6;
      paintMetaLine(line, { size: 9, align: 'right' });
    });
  }

  if (isFormXIXKA && matrix.formXIXKABottomSection) {
    const section = matrix.formXIXKABottomSection;
    const labels = Array.isArray(section.labels) ? section.labels : [];
    const values = Array.isArray(section.values) ? section.values : [];
    const colW = usableWidth / 3;
    const paintKaFooterCells = (cells, { bold = false, size = 8 } = {}) => {
      doc.setFont('helvetica', bold ? 'bold' : 'normal');
      doc.setFontSize(size);
      doc.setTextColor(0, 0, 0);
      const wrapped = [0, 1, 2].map((i) =>
        doc.splitTextToSize(String(cells[i] || ''), Math.max(colW - 8, 24))
      );
      const h = Math.max(14, Math.max(...wrapped.map((lines) => lines.length)) * (size + 2) + 4);
      if (y + h > pageHeight - marginBottom) {
        doc.addPage();
        y = marginTop;
      }
      wrapped.forEach((lines, i) => {
        if (!lines.length) return;
        const cx = marginX + i * colW;
        const textY = y + size + 2;
        if (i === 0) doc.text(lines, cx, textY);
        else if (i === 2) doc.text(lines, cx + colW, textY, { align: 'right' });
        else doc.text(lines, cx + colW / 2, textY, { align: 'center' });
      });
      y += h;
    };
    y += 8;
    paintKaFooterCells(labels, { size: 8 });
    if (values.some((v) => String(v || '').trim())) {
      paintKaFooterCells(values, { size: 9 });
    }
    y += 8;
    paintMetaLine(
      section.signatureText || 'Signature of the contractor or his Representative',
      { size: 8, align: 'right' }
    );
  }

  // Form XIV Employment Card: signature sits under the workman box, not inside it.
  if (isFormXIVEmploymentCardSheet && Array.isArray(matrix.formXIVFooterLines)) {
    matrix.formXIVFooterLines.forEach((line) => {
      const text = String(line || '').trim();
      if (!text) return;
      y += 10;
      paintMetaLine(text, { size: 9, align: 'right' });
    });
  }

  // Also catch a note that landed in the meta band or above the table start.
  if (
    !pendingSystemNote ||
    !pendingUnpaidFootnote ||
    !pendingWhereverApplicableFootnote ||
    pendingForm14RjFootnotes.length < 2 ||
    pendingFormDRJFootnotes.length < FORM_D_RJ_OUTER_FOOTNOTES.length ||
    pendingFormDGJFootnotes.length < FORM_DGJ_GJ_OUTER_FOOTNOTES.length ||
    pendingFormAGJFootnotes.length < FORM_AGJ_GJ_OUTER_FOOTNOTES.length
  ) {
    for (let r = 0; r < rows.length; r += 1) {
      if (!pendingSystemNote && isSystemGeneratedDocumentNoteRow(rows[r])) {
        pendingSystemNote = true;
      }
      if (!pendingUnpaidFootnote) {
        const fn = extractUnpaidAccumulationsFootnoteText(rows[r]);
        if (fn) pendingUnpaidFootnote = fn;
      }
      if (!pendingWhereverApplicableFootnote) {
        const whereverFn = extractFormCWhereverApplicableFootnoteText(rows[r]);
        if (whereverFn) pendingWhereverApplicableFootnote = whereverFn;
      }
      if (isForm14RajasthanFootnoteRow(rows[r]) || isForm14RajasthanFootnoteText(rows[r]?.[0])) {
        const fn = extractForm14RajasthanFootnoteText(rows[r]) || String(rows[r]?.[0] || '').trim();
        if (fn) pushForm14RjFootnote(fn);
      }
      if (
        looksLikeFormDRJSheet &&
        (isFormDRJOuterFootnoteRow(rows[r]) || isFormDRJOuterFootnoteText(rows[r]?.[0]))
      ) {
        pushFormDRJFootnote(rows[r]);
      }
      if (
        !looksLikeFormAGJSheet &&
        (isFormDGJGujaratOuterFootnoteRow(rows[r]) ||
          isFormDGJGujaratOuterFootnoteText(rows[r]?.[0]))
      ) {
        pushFormDGJFootnote(rows[r]);
      }
      if (
        !looksLikeFormBGJSheet &&
        (isFormAGJGujaratOuterFootnoteRow(rows[r]) ||
          isFormAGJGujaratOuterFootnoteText(rows[r]?.[0]))
      ) {
        pushFormAGJFootnote(rows[r]);
      }
    }
  }

  // Form C always shows the unpaid-accumulations definition before the system note.
  if (looksLikeFormCLwfSheet && !pendingUnpaidFootnote) {
    pendingUnpaidFootnote = FORM_C_UNPAID_ACCUMULATIONS_FOOTNOTE;
  }
  if (looksLikeFormCLwfSheet) {
    pendingSystemNote = true;
  }
  if (looksLikeFormCRajasthanSheet && !looksLikeFormCGJSheet) {
    pendingUnpaidFootnote = FORM_C_RJ_FOOTNOTE;
    pendingSystemNote = true;
  }
  // Form C GJ: always keep "*Wherever applicable" outside the table box; never RJ damage note.
  if (looksLikeFormCGJSheet) {
    if (!pendingWhereverApplicableFootnote) {
      pendingWhereverApplicableFootnote = FORM_C_GJ_WHEREVER_APPLICABLE_FOOTNOTE;
    }
    if (isFormCDamageLossFineFootnoteText(pendingUnpaidFootnote)) {
      pendingUnpaidFootnote = '';
    }
    pendingSystemNote = true;
  }

  // Form 14 RJ always keeps both legal notes under the register, above the system note.
  if (looksLikeForm14RjSheet) {
    if (!pendingForm14RjFootnotes.some((t) => isForm14RajasthanOvertimeFootnoteText(t))) {
      pushForm14RjFootnote(FORM_14_RJ_FOOTER_NOTE);
    }
    if (!pendingForm14RjFootnotes.some((t) => isForm14RajasthanDayEntriesNoteText(t))) {
      pushForm14RjFootnote(FORM_14_RJ_DAY_ENTRIES_NOTE);
    }
    pendingSystemNote = true;
  }

  // Form D RJ: always keep Mines / absence-code / E-Form notes outside the attendance box.
  if (looksLikeFormDRJSheet) {
    FORM_D_RJ_OUTER_FOOTNOTES.forEach((note) => pushFormDRJFootnote(note));
    pendingSystemNote = true;
  }

  // Form D GJ: always keep electronic-format / Governor notes outside the muster-roll box.
  if (looksLikeFormDGJSheet) {
    FORM_DGJ_GJ_OUTER_FOOTNOTES.forEach((note) => pushFormDGJFootnote(note));
    pendingSystemNote = true;
  }

  // Form A GJ: always keep age / skill / electronic / wherever notes outside the employee box.
  if (looksLikeFormAGJSheet && !looksLikeFormBGJSheet) {
    FORM_AGJ_GJ_OUTER_FOOTNOTES.forEach((note) => pushFormAGJFootnote(note));
    pendingSystemNote = true;
    // Prefer Form A ***Wherever applicable over Form C *Wherever applicable.
    pendingWhereverApplicableFootnote = '';
  }

  // Form B GJ Excel model: system-generated note only (no Form A footnotes).
  if (looksLikeFormBGJSheet) {
    pendingFormAGJFootnotes.length = 0;
    pendingWhereverApplicableFootnote = '';
    pendingSystemNote = true;
  }

  if (pendingWhereverApplicableFootnote) {
    y += 8;
    paintMetaLine(pendingWhereverApplicableFootnote, { bold: false, size: 8, align: 'left' });
  }
  if (pendingUnpaidFootnote) {
    y += pendingWhereverApplicableFootnote ? 4 : 8;
    paintMetaLine(pendingUnpaidFootnote, { bold: false, size: 8, align: 'left' });
  }
  if (pendingForm14RjFootnotes.length) {
    y += pendingUnpaidFootnote || pendingWhereverApplicableFootnote ? 4 : 8;
    pendingForm14RjFootnotes.forEach((fn, idx) => {
      if (idx > 0) y += 2;
      paintMetaLine(fn, { bold: false, size: 8, align: 'left' });
    });
  }
  if (pendingFormDRJFootnotes.length) {
    y +=
      pendingUnpaidFootnote ||
      pendingWhereverApplicableFootnote ||
      pendingForm14RjFootnotes.length
        ? 4
        : 8;
    const ordered = FORM_D_RJ_OUTER_FOOTNOTES.filter((canon) =>
      pendingFormDRJFootnotes.some(
        (t) =>
          normalizeFormDRJOuterFootnoteText(t).toLowerCase() ===
          normalizeFormDRJOuterFootnoteText(canon).toLowerCase()
      )
    );
    const extras = pendingFormDRJFootnotes.filter(
      (t) =>
        !FORM_D_RJ_OUTER_FOOTNOTES.some(
          (canon) =>
            normalizeFormDRJOuterFootnoteText(t).toLowerCase() ===
            normalizeFormDRJOuterFootnoteText(canon).toLowerCase()
        )
    );
    [...ordered, ...extras].forEach((fn, idx) => {
      if (idx > 0) y += 2;
      paintMetaLine(fn, { bold: false, size: 8, align: 'left' });
    });
  }
  if (pendingFormDGJFootnotes.length) {
    y +=
      pendingUnpaidFootnote ||
      pendingWhereverApplicableFootnote ||
      pendingForm14RjFootnotes.length ||
      pendingFormDRJFootnotes.length
        ? 4
        : 8;
    // Keep * / ** / Governor order stable.
    const ordered = FORM_DGJ_GJ_OUTER_FOOTNOTES.filter((canon) =>
      pendingFormDGJFootnotes.some(
        (t) => normalizeFormDGJFootnoteCompare(t) === normalizeFormDGJFootnoteCompare(canon)
      )
    );
    const extras = pendingFormDGJFootnotes.filter(
      (t) =>
        !FORM_DGJ_GJ_OUTER_FOOTNOTES.some(
          (canon) => normalizeFormDGJFootnoteCompare(t) === normalizeFormDGJFootnoteCompare(canon)
        )
    );
    [...ordered, ...extras].forEach((fn, idx) => {
      if (idx > 0) y += 2;
      paintMetaLine(fn, { bold: false, size: 8, align: 'left' });
    });
  }
  if (pendingFormAGJFootnotes.length) {
    y +=
      pendingUnpaidFootnote ||
      pendingWhereverApplicableFootnote ||
      pendingForm14RjFootnotes.length ||
      pendingFormDRJFootnotes.length ||
      pendingFormDGJFootnotes.length
        ? 4
        : 8;
    const ordered = FORM_AGJ_GJ_OUTER_FOOTNOTES.filter((canon) =>
      pendingFormAGJFootnotes.some(
        (t) => normalizeFormAGJFootnoteCompare(t) === normalizeFormAGJFootnoteCompare(canon)
      )
    );
    const extras = pendingFormAGJFootnotes.filter(
      (t) =>
        !FORM_AGJ_GJ_OUTER_FOOTNOTES.some(
          (canon) => normalizeFormAGJFootnoteCompare(t) === normalizeFormAGJFootnoteCompare(canon)
        )
    );
    [...ordered, ...extras].forEach((fn, idx) => {
      if (idx > 0) y += 2;
      paintMetaLine(fn, { bold: false, size: looksLikeFormAGJSheet ? 9 : 8, align: 'left' });
    });
  }
  if (pendingSystemNote) {
    y +=
      pendingUnpaidFootnote ||
      pendingWhereverApplicableFootnote ||
      pendingForm14RjFootnotes.length ||
      pendingFormDRJFootnotes.length ||
      pendingFormDGJFootnotes.length ||
      pendingFormAGJFootnotes.length
        ? 4
        : 8;
    paintMetaLine(SYSTEM_GENERATED_DOCUMENT_NOTE, {
      bold: false,
      size: looksLikeFormBGJSheet ? 10 : 9,
      align: 'center',
    });
  }

  return y + 12;
};

/**
 * Convert statutory draft Excel/ZIP → PDF with all sheet data (attendance/payroll/leave
 * values included in the draft), headings shown once.
 */
export async function buildStatutoryDraftPdfBlob({
  arrayBuffer,
  fileName = 'draft.pdf',
  title = 'Statutory Draft',
  monthLabel = ''
} = {}) {
  void monthLabel;
  const excelFiles = await collectExcelBuffersFromDraft(arrayBuffer, fileName);
  if (!excelFiles.length) {
    throw new Error('No Excel data found in the draft file to convert to PDF.');
  }

  try {
    const smartBlob = await convertExcelFilesToSmartbrowzPdf({
      excelFiles,
      title,
      fileName
    });
    if (smartBlob && smartBlob.size > 80) {
      return smartBlob;
    }
  } catch (smartErr) {
    console.warn('SmartBrowz PDF conversion failed, using local jsPDF:', smartErr);
  }

  const allMatrices = [];
  for (const file of excelFiles) {
    try {
      const matrices = await collectWorkbookMatrices(file.arrayBuffer, file.label);
      matrices.forEach((m) => {
        // Prefer draft Excel / ZIP entry name so Form_XIX_GJ is visible to detectors
        // even when the worksheet tab is still "Sheet1".
        if (file.label && (!m.name || m.name === 'Sheet1' || looksLikeExcelSheetTabName(m.name))) {
          m.name = file.label;
        }
        m.fileName = file.label || fileName || m.fileName || '';
        // Form XI RJ: drop empty column A left of Serial No. (ZIP entries are
        // named by employee — use parent draft/file name as an extra hint).
        if (
          looksLikeFormXIRajasthanServiceCertificatePdfContext(
            m.metaLines,
            m.rows,
            m.name || m.fileName || fileName
          ) ||
          /form[\s._-]*xi[\s._-]*rj|form_xi_rj/i.test(String(fileName || ''))
        ) {
          const trimmedLead = trimFormXIRajasthanLeadingBlankPdfColumns(
            m.rows,
            m.colCount,
            m.tableStartRow || 0
          );
          m.rows = sanitizeFormXIRajasthanServiceCertificatePdfDataRows(
            trimmedLead.rows,
            trimmedLead.colCount,
            m.tableStartRow || 0
          );
          m.colCount = trimmedLead.colCount;
        }
        // Form L GJ: drop the empty template column before Sr. No. (file name may be
        // the only Gujarat hint when the worksheet tab is still Sheet1).
        if (
          looksLikeFormLGJGujaratPdfContext(
            m.metaLines,
            m.rows,
            m.name,
            m.fileName
          )
        ) {
          const trimmedLead = trimFormLGJGujaratLeadingBlankPdfColumns(
            m.rows,
            m.colCount,
            m.tableStartRow || 0
          );
          m.rows = sanitizeFormLGJGujaratPdfHeaderRows(
            trimmedLead.rows,
            trimmedLead.colCount,
            m.tableStartRow || 0
          );
          m.colCount = trimmedLead.colCount;
        }
        // Form O GJ: drop the empty template column before Sr. No. (file name may be
        // the only Gujarat hint when the worksheet tab is still Sheet1).
        if (
          looksLikeFormOGJGujaratPdfContext(
            m.metaLines,
            m.rows,
            m.name,
            m.fileName
          )
        ) {
          const trimmedLead = trimFormOGJGujaratLeadingBlankPdfColumns(
            m.rows,
            m.colCount,
            m.tableStartRow || 0
          );
          m.rows = trimmedLead.rows;
          m.colCount = trimmedLead.colCount;
        }
        // Form B TN: Fine stays blank; Other deductions stay under Other.
        if (
          looksLikeFormBTamilNaduPdfContext(
            m.metaLines,
            m.rows,
            m.name,
            m.fileName
          )
        ) {
          const normalized = normalizeFormBTamilNaduPdfMatrix(
            m.rows,
            m.colCount,
            m.tableStartRow || 0,
            m.metaLines
          );
          m.rows = normalized.rows;
          m.colCount = normalized.colCount;
          m.tableStartRow = normalized.tableStartRow;
          if (Array.isArray(normalized.metaLines)) m.metaLines = normalized.metaLines;
        }
        // Karnataka XIX: official dotted Excel → boxed wage-slip model (image 1).
        if (
          looksLikeFormXIXKarnatakaPdfContext(m.metaLines, m.rows, m.name, m.fileName) &&
          m.formXIXKarnatakaLayout !== true
        ) {
          applyFormXIXKarnatakaPdfNormalization(
            m,
            normalizeFormXIXKarnatakaWageSlipPdfMatrix(
              m.rows,
              m.colCount,
              m.tableStartRow || 0,
              m.metaLines
            )
          );
        }
        // Gujarat XIX: force normalize using file name hint (Sheet1 alone is not enough).
        if (
          looksLikeFormXIXGJWageSlipPdfContext(m.metaLines, m.rows, m.name, m.fileName) &&
          m.formXIXGJBoxedLayout !== true
        ) {
          const normalized = normalizeFormXIXGJWageSlipPdfMatrix(
            m.rows,
            m.colCount,
            m.tableStartRow || 0,
            m.metaLines
          );
          m.rows = normalized.rows;
          m.colCount = normalized.colCount;
          m.tableStartRow = normalized.tableStartRow;
          m.formXIXAPLayout = normalized.formXIXAPLayout === true;
          m.formXIXAPFooterLines = normalized.formXIXAPFooterLines || [];
          m.formXIXGJHeaderLines = Array.isArray(normalized.formXIXGJHeaderLines)
            ? normalized.formXIXGJHeaderLines
            : [];
          m.formXIXGJLayout = normalized.formXIXGJLayout === true;
          m.formXIXGJBoxedLayout = normalized.formXIXGJBoxedLayout === true;
          if (Array.isArray(normalized.metaLines)) m.metaLines = normalized.metaLines;
        }
        // Karnataka Form XIV: official dotted fill-in Excel → 2-column Employment Card.
        if (
          looksLikeFormXIVKarnatakaPdfContext(m.metaLines, m.rows, m.name, m.fileName) &&
          m.formXIVKALayout !== true
        ) {
          applyFormXIVKarnatakaPdfNormalization(
            m,
            normalizeFormXIVKarnatakaPdfMatrix(
              m.rows,
              m.colCount,
              m.tableStartRow || 0,
              m.metaLines
            )
          );
        }
        // Karnataka Form T: combined muster — never Form XIII Register of Workmen.
        if (
          looksLikeFormTSEKarnatakaPdfContext(m.metaLines, m.rows, m.name, m.fileName) &&
          m.formTKALayout !== true
        ) {
          applyFormTSEKarnatakaPdfNormalization(
            m,
            normalizeFormTSEKarnatakaPdfMatrix(
              m.rows,
              m.colCount,
              m.tableStartRow || 0,
              m.metaLines
            )
          );
        }
        // Karnataka Form Q appointment order — 2-column Excel label|value.
        if (
          looksLikeFormQKarnatakaPdfContext(m.metaLines, m.rows, m.name, m.fileName) &&
          m.formQKALayout !== true
        ) {
          applyFormQKarnatakaPdfNormalization(
            m,
            normalizeFormQKarnatakaPdfMatrix(
              m.rows,
              m.colCount,
              m.tableStartRow || 0,
              m.metaLines
            )
          );
        }
        // Gujarat Form M identity card: sparse Excel → official card layout.
        if (
          looksLikeFormMGJGujaratPdfContext(m.metaLines, m.rows, m.name, m.fileName) &&
          m.formMGJLayout !== true
        ) {
          applyFormMGJGujaratPdfNormalization(
            m,
            normalizeFormMGJGujaratPdfMatrix(
              m.rows,
              m.colCount,
              m.tableStartRow || 0,
              m.metaLines
            )
          );
        }
        // Gujarat Form N Leave Book: identity + accumulation / festival / casual sections.
        if (
          looksLikeFormNGJGujaratPdfContext(m.metaLines, m.rows, m.name, m.fileName) &&
          m.formNGJLayout !== true
        ) {
          applyFormNGJGujaratPdfNormalization(
            m,
            normalizeFormNGJGujaratPdfMatrix(
              m.rows,
              m.colCount,
              m.tableStartRow || 0,
              m.metaLines
            )
          );
        }
        // Karnataka Form F leave register — identity + PART I + PART II Excel model.
        if (
          looksLikeFormFKarnatakaPdfContext(m.metaLines, m.rows, m.name, m.fileName) &&
          m.formFKALayout !== true
        ) {
          applyFormFKarnatakaPdfNormalization(
            m,
            normalizeFormFKarnatakaPdfMatrix(
              m.rows,
              m.colCount,
              m.tableStartRow || 0,
              m.metaLines
            )
          );
        }
        if (
          looksLikeFormPKarnatakaPdfContext(m.metaLines, m.rows, m.name, m.fileName) &&
          m.formPKALayout !== true
        ) {
          applyFormPKarnatakaPdfNormalization(
            m,
            normalizeFormPKarnatakaPdfMatrix(
              m.rows,
              m.colCount,
              m.tableStartRow || 0,
              m.metaLines,
              m.fileName || m.name || ''
            )
          );
        }
        allMatrices.push(m);
      });
    } catch (err) {
      console.warn('Skipped draft Excel for PDF:', file.label, err);
    }
  }

  let matricesForPdf = filterStatutoryPdfMatrices(allMatrices);
  if (
    matricesForPdf.some((m) =>
      looksLikeFormXVAPServiceCertificatePdfContext(m.metaLines, m.rows, m.name)
    )
  ) {
    const withWorkman = matricesForPdf.filter((m) => {
      if (!looksLikeFormXVAPServiceCertificatePdfContext(m.metaLines, m.rows, m.name)) return true;
      const blob = [...(m.metaLines || []), ...(m.rows || []).slice(0, 40).flat()].join('\n');
      const match = blob.match(/name\s+and\s+address\s+of\s+the\s+workm[ae]n[:\s.]+([^\n]+)/i);
      const value = String(match?.[1] || '')
        .replace(/^[:.\s]+/, '')
        .trim();
      if (!value || /^enter\b/i.test(value)) return false;
      return true;
    });
    if (withWorkman.length) matricesForPdf = withWorkman;
  }
  if (!matricesForPdf.length) {
    throw new Error('Draft Excel has no readable rows to put in the PDF.');
  }

  const maxCols = Math.max(...matricesForPdf.map((m) => m.colCount));
  const anyFormW = matricesForPdf.some((m) =>
    looksLikeFormWPdfContext(m.metaLines, m.rows, m.tableStartRow || 0)
  );
  const anyFormTKA = matricesForPdf.some(
    (m) =>
      m.formTKALayout === true ||
      looksLikeFormTSEKarnatakaPdfContext(m.metaLines, m.rows, m.name, m.fileName)
  );
  const anyAccidentBook = matricesForPdf.some((m) => {
    const blob = [...(m.metaLines || []), ...(m.rows || []).slice(0, 8).flat(), m.name || '']
      .join(' ')
      .toLowerCase();
    return /accident\s+book|form\s*(?:no\.?\s*)?11\b/.test(blob);
  });
  const anyForm25 = matricesForPdf.some((m) =>
    looksLikeForm25TamilNaduPdfContext(m.metaLines, m.rows, m.name)
  );
  const anyForm14Rj = matricesForPdf.some((m) =>
    looksLikeForm14RajasthanPdfContext(m.metaLines, m.rows, m.name)
  );
  const anyFormXIXWageSlip = matricesForPdf.some((m) =>
    looksLikeFormXIXWageSlipPdfContext(m.metaLines, m.rows, m.name)
  );
  const anyFormXIVEmploymentCard = matricesForPdf.some((m) =>
    looksLikeFormXIVEmploymentCardPdfContext(m.metaLines, m.rows, m.name)
  );
  const anyFormMGJIdentityCard = matricesForPdf.some(
    (m) =>
      m.formMGJLayout === true ||
      looksLikeFormMGJGujaratPdfContext(m.metaLines, m.rows, m.name, m.fileName)
  );
  const anyFormQKAAppointment = matricesForPdf.some(
    (m) =>
      m.formQKALayout === true ||
      looksLikeFormQKarnatakaPdfContext(m.metaLines, m.rows, m.name, m.fileName)
  );
  const anyFormNGJLeaveBook = matricesForPdf.some(
    (m) =>
      m.formNGJLayout === true ||
      looksLikeFormNGJGujaratPdfContext(m.metaLines, m.rows, m.name, m.fileName)
  );
  const anyFormFKALeaveRegister = matricesForPdf.some(
    (m) =>
      m.formFKALayout === true ||
      looksLikeFormFKarnatakaPdfContext(m.metaLines, m.rows, m.name, m.fileName)
  );
  const anyFormPKANotice = matricesForPdf.some(
    (m) =>
      m.formPKALayout === true ||
      looksLikeFormPKarnatakaPdfContext(m.metaLines, m.rows, m.name, m.fileName)
  );
  const anyFormPGJ = matricesForPdf.some((m) =>
    looksLikeFormPGJGujaratPdfContext(m.metaLines, m.rows, m.name, m.fileName)
  );
  const anyFormQMaharashtra = matricesForPdf.some((m) =>
    looksLikeFormQMaharashtraPdfContext(m.metaLines, m.rows, m.name, m.fileName)
  );
  // Form 14 RJ has few columns but very long header labels — landscape keeps them aligned.
  // Form XIX Wage Slip / Form XIV Employment Card / Form M GJ / Form Q KA stay portrait A4.
  const forcePortraitCard =
    anyFormXIXWageSlip ||
    anyFormXIVEmploymentCard ||
    anyFormMGJIdentityCard ||
    anyFormQKAAppointment ||
    anyFormPKANotice;
  const wide =
    (!forcePortraitCard && maxCols > 8) ||
    anyFormW ||
    anyFormTKA ||
    anyAccidentBook ||
    anyForm25 ||
    anyForm14Rj ||
    anyFormNGJLeaveBook ||
    anyFormFKALeaveRegister ||
    anyFormPGJ ||
    anyFormQMaharashtra;
  // Form W (~30 wage/deduction cols) and Form T KA (~40 identity/attendance/wage cols)
  // need A2 landscape so amounts stay on one line.
  // Form P GJ / Form Q MH (~60 worker/day/wage cols) match the Excel muster-roll on A2 landscape.
  // Form 11 Accident Book (~18 cols with long headers) and other wide registers need A3.
  const veryWide =
    anyFormW ||
    anyFormTKA ||
    anyFormPGJ ||
    anyFormQMaharashtra ||
    anyAccidentBook ||
    anyForm25 ||
    anyFormNGJLeaveBook ||
    (!forcePortraitCard && maxCols > 14);
  const doc = new jsPDF({
    unit: 'pt',
    format: forcePortraitCard
      ? 'a4'
      : anyFormW || anyFormTKA || anyFormPGJ || anyFormQMaharashtra
        ? 'a2'
        : veryWide
          ? 'a3'
          : 'a4',
    orientation: forcePortraitCard ? 'portrait' : wide ? 'landscape' : 'portrait'
  });

  // Bordered header draws form name for every sheet — skip a floating duplicate title.
  const firstMatrix = matricesForPdf[0];
  const pdfHeaderOpts = { preferredTitle: title, fileName };
  const firstHeaderModel = buildStatutoryPdfHeaderModel(
    firstMatrix?.metaLines || [],
    firstMatrix?.rows || [],
    firstMatrix?.tableStartRow || 0,
    firstMatrix?.name || '',
    pdfHeaderOpts
  );
  const hasBorderedHeader =
    (firstHeaderModel.titles || []).length > 0 || (firstHeaderModel.fields || []).length > 0;
  const isApPdf =
    firstHeaderModel.formXXAP === true ||
    firstHeaderModel.formXXIAP === true ||
    firstHeaderModel.formXVIIAP === true ||
    firstHeaderModel.formXVIAP === true ||
    firstHeaderModel.formXIIIAP === true ||
    firstHeaderModel.formXXIIIAPSE === true;
  // Use a real form title only — never the Excel/ZIP draft file name as a PDF heading.
  const headingCandidate = String(title || '')
    .replace(EXCEL_EXT_RE, '')
    .replace(/\.zip$/i, '')
    .trim()
    .slice(0, 120);
  const heading =
    headingCandidate && !looksLikeExcelDraftFileLabel(headingCandidate) ? headingCandidate : '';
  if (!hasBorderedHeader && heading) {
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(isApPdf ? 9 : 11);
    doc.setTextColor(0, 0, 0);
    doc.text(heading, doc.internal.pageSize.getWidth() / 2, 20, { align: 'center' });
  }

  let y = hasBorderedHeader || !heading ? 20 : 32;
  for (let i = 0; i < matricesForPdf.length; i += 1) {
    if (i > 0) {
      doc.addPage();
      y = 28;
    }
    y = drawMatrixSheet(doc, matricesForPdf[i], y, pdfHeaderOpts);
  }

  return doc.output('blob');
}

export async function buildFormLGJDraftPdfBlob(opts = {}) {
  return buildStatutoryDraftPdfBlob(opts);
}

export { buildFormXVAPServiceCertificatePdfBlob };

export function draftFileNameToPdfName(fileName) {
  const base = String(fileName || 'statutory-draft')
    .replace(/\.(xlsx|xls|xlsm|xlsb|zip)$/i, '')
    .replace(/[^\w.-]+/g, '_')
    .replace(/_+/g, '_')
    .replace(/^_|_$/g, '');
  return `${base || 'statutory-draft'}.pdf`;
}

/** Test-only helpers for Form W PDF layout detection. */
export const statutoryDraftPdfTestUtils = {
  isWageRegisterColHeaderBlob,
  isFormCLwfColHeaderBlob,
  isForm14RajasthanColHeaderBlob,
  isFormBRajasthanWageRateGroupLabel,
  looksLikeFormBRajasthanPdfContext,
  looksLikeForm14RajasthanPdfContext,
  looksLikeFormDRajasthanPdfContext,
  looksLikeForm11RajasthanPdfContext,
  looksLikeFormXIXRajasthanOvertimePdfContext,
  looksLikeFormXIXWageSlipPdfContext,
  looksLikeFormXRajasthanEmploymentCardPdfContext,
  buildFormXRajasthanEmploymentHeaderModel,
  looksLikeFormXIRajasthanServiceCertificatePdfContext,
  buildFormXIRajasthanServiceHeaderModel,
  looksLikeFormXVRajasthanWageSlipPdfContext,
  buildFormXVRajasthanWageSlipHeaderModel,
  isFormXIXWageSlipColHeaderBlob,
  isFormXIXAPWageSlipPdfContext,
  normalizeFormXIXAPWageSlipPdfMatrix,
  normalizeFormXIXWageSlipPdfMatrix,
  looksLikeFormXIXGJWageSlipPdfContext,
  normalizeFormXIXGJWageSlipPdfMatrix,
  looksLikeFormXIXKarnatakaPdfContext,
  normalizeFormXIXKarnatakaWageSlipPdfMatrix,
  looksLikeFormXIXTamilNaduPdfContext,
  resolveFormXIXTamilNaduNetAmountPdfSpan,
  looksLikeFormBTamilNaduPdfContext,
  rewriteFormBTamilNaduPdfHeader,
  isFormBTamilNaduPdfSignatoryRow,
  isFormBTamilNaduPdfSignatoryText,
  rewriteFormCTamilNaduLwfPdfHeader,
  isFormCTamilNaduLwfPdfSignatoryRow,
  isFormCTamilNaduLwfPdfSignatoryText,
  detectFormBTamilNaduPdfGroupBands,
  isFormBTamilNaduAmountsDeductedGroupLabel,
  normalizeFormBTamilNaduPdfMatrix,
  looksLikeFormXIVEmploymentCardPdfContext,
  normalizeFormXIVEmploymentCardPdfMatrix,
  looksLikeFormXIVKarnatakaPdfContext,
  normalizeFormXIVKarnatakaPdfMatrix,
  looksLikeFormTSEKarnatakaPdfContext,
  normalizeFormTSEKarnatakaPdfMatrix,
  getFormTSEKarnatakaHeaderTitles,
  extractFormTSEKarnatakaHeaderFields,
  looksLikeFormQKarnatakaPdfContext,
  normalizeFormQKarnatakaPdfMatrix,
  looksLikeFormMGJGujaratPdfContext,
  normalizeFormMGJGujaratPdfMatrix,
  looksLikeFormNGJGujaratPdfContext,
  normalizeFormNGJGujaratPdfMatrix,
  looksLikeFormFKarnatakaPdfContext,
  normalizeFormFKarnatakaPdfMatrix,
  looksLikeFormPKarnatakaPdfContext,
  normalizeFormPKarnatakaPdfMatrix,
  isFormXIVEmploymentCardWorkmanLabelText,
  scrubForm11RajasthanTotalHoursPdfColumn,
  trimForm14RajasthanLeadingBlankPdfColumns,
  trimFormXIXRajasthanLeadingBlankPdfColumns,
  trimFormXIRajasthanLeadingBlankPdfColumns,
  sanitizeFormXIRajasthanServiceCertificatePdfDataRows,
  trimFormXRajasthanEmploymentLeadingBlankPdfColumns,
  sanitizeFormXRajasthanEmploymentPdfDataRows,
  isFormXRajasthanEmploymentNumericWageRate,
  isForm25TamilNaduColHeaderBlob,
  looksLikeForm25TamilNaduPdfContext,
  findRemarksColumnIndex,
  trimTrailingBlankPdfColumns,
  isForm25TamilNaduEmployeePdfRow,
  trimForm25TamilNaduPdfTrailingEmployeeRows,
  looksLikeAuxiliaryOrPivotPdfSheet,
  looksLikeForm15TamilNaduPdfContext,
  looksLikeForm15Part1PreferredContext,
  rewriteForm15Part1PdfTitles,
  isStandaloneFormXTitle,
  detectLeaveCategoryBands,
  detectStatutoryGroupHeaderBands,
  isLeaveCategoryGroupLabel,
  isWageDeductionGroupLabel,
  isFormVIFestivalGroupLabel,
  isStatutoryGroupHeaderLabel,
  filterStatutoryPdfMatrices,
  stripLeakedPivotRowsFromPdfMatrix,
  isFormWAdminBandBlob,
  isFormWAdminValueRow,
  looksLikeFormWPdfContext,
  formWTamilNaduColumnWeight,
  formXXVIITamilNaduColumnWeight,
  formXXVITamilNaduColumnWeight,
  formXXVITamilNaduLongestHeaderWord,
  isFormXXVITamilNaduTrailingWideHeader,
  wrapPdfTextPreferWholeWords,
  redistributeFormXXVIPdfColumnWidths,
  looksLikeFormXXVIITamilNaduRegisterPdfContext,
  looksLikeFormXVIIITamilNaduPdfContext,
  FORM_XVIII_TN_PDF_TABLE_FONT_SIZE,
  coalesceFormXVIIITamilNaduTitleLines,
  detectFormXVIIITamilNaduPdfGroupBands,
  isFormXVIIITamilNaduWagesCumMusterTitle,
  isFormXVIIITamilNaduAmountOfWagesEarnedGroupLabel,
  extractMonthFromWagePeriodLine,
  isPurePdfNumericText,
  isPureNilPdfText,
  isNilOfTheMonthPdfText,
  resolveNilOfTheMonthPdfSpan,
  isUnpaidAccumulationsFootnoteText,
  isUnpaidAccumulationsFootnoteRow,
  FORM_C_UNPAID_ACCUMULATIONS_FOOTNOTE,
  FORM_C_RJ_FOOTNOTE,
  FORM_C_GJ_WHEREVER_APPLICABLE_FOOTNOTE,
  FORM_DGJ_GJ_ELECTRONIC_FORMAT_NOTE_1,
  FORM_DGJ_GJ_ELECTRONIC_FORMAT_NOTE_2,
  FORM_DGJ_GJ_GOVERNOR_ORDER_NOTE,
  FORM_DGJ_GJ_OUTER_FOOTNOTES,
  FORM_AGJ_GJ_AGE_NOTE,
  FORM_AGJ_GJ_SKILL_NOTE,
  FORM_AGJ_GJ_ELECTRONIC_NOTE,
  FORM_AGJ_GJ_WHEREVER_NOTE,
  FORM_AGJ_GJ_OUTER_FOOTNOTES,
  looksLikeFormCRajasthanPdfContext,
  looksLikeFormCGJLoanRecoveriesPdfContext,
  looksLikeFormDGJMusterRollPdfContext,
  looksLikeFormAGJGujaratPdfContext,
  looksLikeFormAPdfContext,
  looksLikeFormBGJGujaratPdfContext,
  looksLikeFormKGJGujaratPdfContext,
  rewriteFormKGJGujaratPdfHeader,
  trimFormKGJGujaratLeadingBlankPdfColumns,
  looksLikeFormLGJGujaratPdfContext,
  trimFormLGJGujaratLeadingBlankPdfColumns,
  sanitizeFormLGJGujaratPdfHeaderRows,
  detectFormLGJGujaratPdfGroupBands,
  formLGJGujaratColumnWeight,
  looksLikeFormPGJGujaratPdfContext,
  rewriteFormPGJGujaratPdfHeader,
  trimFormPGJGujaratLeadingBlankPdfColumns,
  detectFormPGJGujaratPdfGroupBands,
  formPGJGujaratColumnWeight,
  isFormPGJDateOfMonthGroupLabel,
  normalizeFormPGJGujaratPdfMatrix,
  dropFormPGJGujaratExtraDayPdfColumns,
  blankFormPGJGujaratMinimumRateOfWagesPdfCells,
  resolveFormPGJGujaratPdfDaysInMonth,
  looksLikeFormOGJGujaratPdfContext,
  rewriteFormOGJGujaratPdfHeader,
  trimFormOGJGujaratLeadingBlankPdfColumns,
  detectFormLGJGujaratPdfGroupBands,
  formLGJGujaratColumnWeight,
  extractFormBGJMinimumWagesBox,
  isFormBGJMinimumWagesMetaText,
  normalizeFormBGJGujaratPdfMatrix,
  reorderFormBGJPdfHeaderFields,
  buildFormAGJPdfSplitFieldRows,
  reorderFormAGJPdfHeaderFields,
  reorderFormCGJPdfHeaderFields,
  isStatutoryFieldMetaLine,
  isFormCWhereverApplicableFootnoteText,
  isFormCDamageLossFineFootnoteText,
  isFormDGJGujaratOuterFootnoteText,
  isFormDGJGujaratOuterFootnoteRow,
  isFormAGJGujaratOuterFootnoteText,
  isFormAGJGujaratOuterFootnoteRow,
  isForm14RajasthanOvertimeFootnoteText,
  isForm14RajasthanDayEntriesNoteText,
  isForm14RajasthanFootnoteText,
  isForm14RajasthanFootnoteRow,
  extractForm14RajasthanFootnoteText,
  FORM_14_RJ_FOOTER_NOTE,
  FORM_14_RJ_DAY_ENTRIES_NOTE,
  isFormDRJOuterFootnoteText,
  isFormDRJOuterFootnoteRow,
  extractFormDRJOuterFootnoteText,
  FORM_D_RJ_OUTER_FOOTNOTES,
  FORM_D_RJ_RELAY_MINES_NOTE,
  FORM_D_RJ_ABSENCE_CODES_NOTE,
  FORM_D_RJ_E_FORM_NOTE,
  extractFormWGenderBox,
  sheetToDenseMatrix,
  looksLikeExcelDraftFileLabel,
  looksLikeExcelSheetTabName,
  buildStatutoryPdfHeaderModel,
  isStatutoryTitleMetaLine,
  expandStatutoryMetaSegments,
  normalizeStatutoryMultilineText,
  isPdfSerialNumberHeader,
  statutoryHeaderColumnWeight,
  enforcePdfColumnMinWidths,
  resolveLeafHeaderTexts,
   // Form XX Andhra Pradesh
  isFormXXDateOfRecoveryGroupLabel,
  isFormXXAPAdministrativeRow,
  getFormXXAPHeaderTitles,
  isFormXXAPTableHeaderRow,
  looksLikeFormXXAPPdfContext,
  getFormXXIAPHeaderTitles,
  looksLikeFormXXIAPPdfContext,
  getFormXVIIAPHeaderTitles,
  isFormXVIIAdministrativeRow,
  isFormXVIAPTableHeaderRow,
  looksLikeFormXVIIAPPdfContext,
  getFormXVIAPHeaderTitles,
  isFormXVIIDateNumberRow,
  isFormXVIITableHeaderRow,
  looksLikeFormXVIAPPdfContext,
  getFormXIIIAPHeaderTitles,
  uniqueExpandedTitleBands,
  isApPdfHeadingRow,
  looksLikeFormXXIIIAPSEPdfContext,
  getFormXXIIIAPSEHeaderTitles,
  extractFormXXIIIAPSEAdminLayout,
  isFormXXIIIAPSETableHeaderRow,
  isFormXXIIIAPSEAdminRow,
  trimFormXXIIIAPSEEmptyPdfColumns,
};