import { normalizeStatutoryHeaderLabel } from '../../utils/statutorySiteCompanyHeaders';

/**
 * Tamil Nadu Form U template columns (Shops & Establishments register — matches Form_U_-_TamilNadu.xlsx).
 * Order must stay aligned with the Excel header row:
 * DOB → Date of Joining → Designation → Present/Permanent → EPF → ESIC → Aadhaar → … → Bank.
 */
export const FORM_U_TAMILNADU_HEADERS = [
  'S.No',
  'Name of the employee',
  'Worker Identity No.',
  'Gender',
  'Father / Spouse Name',
  'Date of Birth',
  'Date of Joining',
  'Designation',
  'Present Address',
  'Permanent Address',
  "Employee's Provident Fund No.",
  "Employee's State Insurance Corporation No.",
  'Aadhaar No.',
  'Date on which completion of 480 days of service',
  'Date on which made permanent',
  'Period of Suspension if any',
  'Bank A/c Number, Name of Bank, Branch (Indian Financial System Code) (IFSC Code)',
  'Photo',
  'Mobile Number',
  'Email ID',
  'Specimen Signature / Thumb Impression',
  'Date of Exit',
  'Reason for Exit',
  'Remarks',
  // Not in all TN templates — keep after Remarks so it does not shift Photo/Email columns.
  'Bank Address'
];

const normalizeFormUTamilNaduText = (text) =>
  String(text ?? '')
    .replace(/\r?\n/g, ' ')
    .replace(/[_-]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();

function pickFirstNonEmptyString(...values) {
  for (const value of values) {
    if (value == null) continue;
    if (typeof value === 'object') {
      const nested =
        value.zc_display_value ?? value.display_value ?? value.displayValue ?? value.name ?? value.Name;
      const nestedText = String(nested ?? '').trim();
      if (nestedText) return nestedText;
      continue;
    }
    const text = String(value).trim();
    if (text) return text;
  }
  return '';
}

function findEmpValueByNormalizedKey(emp, targetKey) {
  if (!emp || typeof emp !== 'object' || !targetKey) return '';
  const target = String(targetKey).toLowerCase().replace(/[^a-z0-9]/g, '');
  for (const [key, value] of Object.entries(emp)) {
    const norm = String(key || '')
      .toLowerCase()
      .replace(/[^a-z0-9]/g, '');
    if (norm === target) {
      const text = pickFirstNonEmptyString(value);
      if (text) return text;
    }
  }
  return '';
}

/** Filename Form_U_-_TamilNadu.xlsx: word-boundary after "u" fails because "_" is a word char. */
export function looksLikeFormUFilename(text) {
  const s = String(text || '').toLowerCase();
  return /(?:^|[^a-z0-9])form[\s._-]*u(?:[^a-z0-9]|$)/i.test(s) || /\bform\s+u\b/i.test(s);
}

export function looksLikeTamilNaduText(text) {
  const s = String(text || '').toLowerCase();
  return /tamil[\s._-]*nadu|tamilnadu/.test(s);
}

export function buildFormUTamilNaduContextBlob(hints = {}) {
  const item = hints.item || hints;
  return [
    hints.fileName,
    hints.formFileName,
    hints.formHeaderTitle,
    hints.formHeaderSubtitle,
    hints.sheetText,
    item?.formName,
    item?.FormName,
    item?.state,
    item?.State,
    item?.description,
    item?.Description,
    item?.act,
    item?.Act,
    item?.formFileName,
    item?.FormFileName,
    Array.isArray(hints.tableHeaders) ? hints.tableHeaders.join(' ') : ''
  ]
    .filter(Boolean)
    .join(' ')
    .toLowerCase();
}

export function isFormUTamilNaduEmployeeRegisterContext(hints = {}) {
  const blob = buildFormUTamilNaduContextBlob(hints);
  if (!looksLikeFormUFilename(blob) && !/\bform\s+u\b/i.test(blob)) return false;
  if (!looksLikeTamilNaduText(blob)) return false;
  if (
    /confidential\s+character|rule\s*34/.test(blob) &&
    !/employee\s+identification|name\s+of\s+the\s+employee|worker\s+identity/.test(blob)
  ) {
    return false;
  }
  return true;
}

/** True when sheet/modal headers look like TN Form U (EPF/ESIC/480/suspension layout). */
export function headersIndicateFormUTamilNaduRegister(headers) {
  if (!Array.isArray(headers) || headers.length === 0) return false;
  const joined = headers.map((h) => normalizeFormUTamilNaduText(h)).join(' ');
  const hasAddress = /present\s+address/.test(joined) && /permanent\s+address/.test(joined);
  const hasTnColumns =
    /provident\s+fund|employee.?s\s+provident/.test(joined) ||
    /state\s+insurance|esic/.test(joined) ||
    /480\s+days/.test(joined) ||
    /made\s+permanent/.test(joined) ||
    /period\s+of\s+suspension/.test(joined);
  const hasLegacyIdentity =
    (/worker\s+identity/.test(joined) || /employee\s+identification/.test(joined)) &&
    (/photo/.test(joined) || /bank/.test(joined));
  return hasAddress && (hasTnColumns || hasLegacyIdentity);
}

/** Template headers that must export via label match (not FORMAT2 pregnancy/wages order). */
export function headersIndicateFormUTamilNaduTemplateLayout(headers) {
  if (!Array.isArray(headers) || headers.length === 0) return false;
  const joined = headers.map((h) => normalizeFormUTamilNaduText(h)).join(' ');
  const hasAddress = /present\s+address/.test(joined) && /permanent\s+address/.test(joined);
  const hasTnMiddle =
    /provident\s+fund|state\s+insurance|480\s+days|made\s+permanent|period\s+of\s+suspension/.test(
      joined
    );
  const hasTnTail =
    /bank/.test(joined) &&
    (/email|e\s*mail/.test(joined) || /mobile/.test(joined)) &&
    /photo/.test(joined);
  return hasAddress && (hasTnMiddle || hasTnTail);
}

/** Draft/template contamination: header cells contain employee names or VE IDs. */
export function formUTamilNaduHeadersLookContaminated(headers) {
  const list = Array.isArray(headers) ? headers : [];
  return list.some((h) => {
    const s = String(h || '');
    if (/\bVE\d{3,}\b/i.test(s)) return true;
    if (
      /name\s+of\s+the\s+employee/i.test(s) &&
      s.replace(/name\s+of\s+the\s+employee/i, '').trim().split(/\s+/).filter(Boolean).length >= 2
    ) {
      return true;
    }
    if (
      /employee\s+identification/i.test(s) &&
      (/\bVE\d/i.test(s) || s.replace(/employee\s+identification\s*no\.?/i, '').trim().length > 8)
    ) {
      return true;
    }
    return false;
  });
}

function remapTruncatedFormUHeaderLabel(header) {
  const compact = normalizeStatutoryHeaderLabel(header);
  if (!compact) return String(header || '').trim();
  if (
    compact === 'bank a c' ||
    compact === 'bank ac' ||
    compact === 'bank a/c' ||
    (compact.includes('bank') &&
      (compact.includes('a c') || compact.includes('ifsc') || compact.includes('financial system')) &&
      !compact.includes('address'))
  ) {
    return 'Bank A/c Number, Name of Bank, Branch (Indian Financial System Code) (IFSC Code)';
  }
  if (compact === 'bank address' || compact === 'bank addr') {
    return 'Bank Address';
  }
  if (compact === 'e mail id' || compact === 'email id' || compact === 'e mail') {
    return 'Email ID';
  }
  if (compact.includes('aadhaar') || compact.includes('aadhar')) {
    if (compact === 'aadhaar' || compact === 'aadhar') return 'Aadhaar No.';
  }
  return String(header || '').trim();
}

function headerLabelKey(header) {
  return normalizeStatutoryHeaderLabel(remapTruncatedFormUHeaderLabel(header));
}

/**
 * Alias buckets for TN Form U export — must distinguish Permanent Address vs Date made permanent.
 */
export function formUTamilNaduHeaderAliasBucket(header) {
  const norm = normalizeFormUTamilNaduText(header).replace(/[^a-z0-9]+/g, ' ').trim();
  if (!norm) return '';
  if (/^s\.?\s*no|serial number|sl\.?\s*no/.test(norm)) return 'sno';
  if (/name of the (employee|worker)|employee name|worker name|^name$/.test(norm)) return 'name';
  // Email before worker-id patterns so "Email ID" is never treated as an employee id.
  if (/\be\s*mail\b|\bemail\b/.test(norm)) return 'email';
  if (
    /worker\s+identity|employee\s+identification|identify\s+no/.test(norm) ||
    (/\bemployee\s+id\b|\bemp\s*id\b/.test(norm) && !/\bemail\b/.test(norm))
  ) {
    return 'workerid';
  }
  if (/\bgender\b|\bsex\b/.test(norm)) return 'gender';
  if (/father|spouse|husband/.test(norm)) return 'father';
  if (/date of birth|\bdob\b/.test(norm)) return 'dob';
  if (/date of joining|date of entry|date on which joined|\bdoj\b/.test(norm)) return 'doj';
  if (/designation|nature of work/.test(norm)) return 'designation';
  if (/present/.test(norm) && /address/.test(norm)) return 'present_address';
  if (/permanent/.test(norm) && /address/.test(norm)) return 'permanent_address';
  if (/provident|epf|\buan\b/.test(norm)) return 'epf';
  if (/state insurance|esic|\besi\b/.test(norm)) return 'esic';
  if (/aadhaar|aadhar/.test(norm)) return 'aadhaar';
  if (/480/.test(norm)) return '480days';
  if (/made permanent/.test(norm)) return 'made_permanent';
  if (/suspension/.test(norm)) return 'suspension';
  if (/date on which left/.test(norm)) return 'date_left';
  if (/period of service/.test(norm)) return 'period_service';
  if (/bank/.test(norm) && /address/.test(norm)) return 'bank_address';
  if (/bank/.test(norm) && (/account|ifsc|branch|financial|a c|ac no/.test(norm))) return 'bank';
  if (/^photo/.test(norm)) return 'photo';
  if (/mobile|contact|phone/.test(norm)) return 'mobile';
  if (/specimen|signature|thumb|thump/.test(norm)) return 'signature';
  if (/date.*exit|date of exit/.test(norm)) return 'dateexit';
  if (/reason.*exit/.test(norm)) return 'reasonexit';
  if (/^remarks?$/.test(norm)) return 'remarks';
  return norm;
}

/** Prefer template/parsed headers; fill any missing TN columns without replacing order. */
export function normalizeFormUTamilNaduEmployeeRegisterHeaders(headers, hints = {}) {
  const list = (Array.isArray(headers) ? headers : [])
    .map((h) => remapTruncatedFormUHeaderLabel(h))
    .filter(Boolean);
  const canonical = [...FORM_U_TAMILNADU_HEADERS];

  if (list.length === 0 || formUTamilNaduHeadersLookContaminated(list)) {
    return canonical;
  }

  // Keep template order when it is already the TN Form U layout.
  if (headersIndicateFormUTamilNaduTemplateLayout(list)) {
    const seen = new Set();
    const out = [];
    list.forEach((header) => {
      const key = formUTamilNaduHeaderAliasBucket(header) || headerLabelKey(header);
      if (!key || seen.has(key)) return;
      seen.add(key);
      out.push(header);
    });
    canonical.forEach((header) => {
      const key = formUTamilNaduHeaderAliasBucket(header) || headerLabelKey(header);
      if (!key || seen.has(key)) return;
      seen.add(key);
      out.push(header);
    });
    return out.length > 0 ? out : canonical;
  }

  // Contaminated FORMAT2-style modal headers → force TN template columns.
  const seen = new Set();
  const out = [];
  canonical.forEach((header) => {
    const key = formUTamilNaduHeaderAliasBucket(header) || headerLabelKey(header);
    if (!key || seen.has(key)) return;
    seen.add(key);
    out.push(header);
  });
  list.forEach((header) => {
    const key = formUTamilNaduHeaderAliasBucket(header) || headerLabelKey(header);
    if (!key || seen.has(key)) return;
    seen.add(key);
    out.push(header);
  });
  return out.length > 0 ? out : canonical;
}

function buildAddressFromZohoComponents(emp = {}) {
  const addressLine1 = pickFirstNonEmptyString(
    emp.Address_Line_1,
    emp['Address_Line_1'],
    emp.AddressLine1,
    emp['Address Line 1'],
    findEmpValueByNormalizedKey(emp, 'addressline1')
  );
  const addressLine2 = pickFirstNonEmptyString(
    emp.Address_Line_2,
    emp['Address_Line_2'],
    emp.AddressLine2,
    emp['Address Line 2'],
    findEmpValueByNormalizedKey(emp, 'addressline2')
  );
  const city = pickFirstNonEmptyString(
    emp.City,
    emp['City'],
    emp.City1,
    emp['City1'],
    emp.City2,
    emp['City2'],
    findEmpValueByNormalizedKey(emp, 'city')
  );
  const state = pickFirstNonEmptyString(emp.State, emp['State'], findEmpValueByNormalizedKey(emp, 'state'));
  const country = pickFirstNonEmptyString(
    emp.Country,
    emp['Country'],
    emp.Country1,
    emp['Country1'],
    findEmpValueByNormalizedKey(emp, 'country')
  );
  return [addressLine1, addressLine2, city, state, country].filter(Boolean).join(', ');
}

export function resolveFormUTamilNaduPermanentAddress(emp = {}) {
  return pickFirstNonEmptyString(
    emp.Permanent_Address,
    emp['Permanent_Address'],
    emp['Permanent address'],
    emp.PermanentAddress,
    emp['PermanentAddress'],
    emp.permanentAddress,
    emp.Permanentaddress,
    emp['Permanentaddress'],
    findEmpValueByNormalizedKey(emp, 'permanentaddress'),
    findEmpValueByNormalizedKey(emp, 'permanent_address'),
    buildAddressFromZohoComponents(emp)
  );
}

export function resolveFormUTamilNaduPresentAddress(emp = {}) {
  return resolveFormUTamilNaduPermanentAddress(emp);
}

export function resolveFormUTamilNaduBankAddress(emp = {}) {
  return pickFirstNonEmptyString(
    emp.Bank_Address,
    emp['Bank_Address'],
    emp['Bank Address'],
    emp.bankAddress,
    emp.BankAddress,
    emp['BankAddress'],
    findEmpValueByNormalizedKey(emp, 'bankaddress'),
    findEmpValueByNormalizedKey(emp, 'bank_address')
  );
}

export function isFormUTamilNaduBankAddressHeader(header) {
  return formUTamilNaduHeaderAliasBucket(header) === 'bank_address';
}

export function isFormUTamilNaduCombinedBankHeader(header) {
  return formUTamilNaduHeaderAliasBucket(header) === 'bank';
}

export function buildFormUTamilNaduCombinedBankDetails(emp = {}) {
  const accountNumber = pickFirstNonEmptyString(
    emp.Bank_Account_Number,
    emp['Bank_Account_Number'],
    emp.Account_Number,
    emp['Account_Number'],
    emp.AccountNumber,
    emp['Account Number'],
    emp.accountNumber,
    findEmpValueByNormalizedKey(emp, 'bankaccountnumber'),
    findEmpValueByNormalizedKey(emp, 'accountnumber'),
    findEmpValueByNormalizedKey(emp, 'account_number')
  );
  const bankName = pickFirstNonEmptyString(
    emp.Bank_Name,
    emp['Bank_Name'],
    emp.BankName,
    emp['Bank Name'],
    emp.bankName,
    findEmpValueByNormalizedKey(emp, 'bankname'),
    findEmpValueByNormalizedKey(emp, 'bank_name')
  );
  const branchName = pickFirstNonEmptyString(
    emp.Branch_Name,
    emp['Branch_Name'],
    emp.BranchName,
    emp['Branch Name'],
    emp.branchName,
    findEmpValueByNormalizedKey(emp, 'branchname')
  );
  const ifscCode = pickFirstNonEmptyString(
    emp.IFSC_Code,
    emp['IFSC_Code'],
    emp.IFSCCode,
    emp['IFSC Code'],
    emp.ifscCode,
    emp.IFSC,
    emp['IFSC'],
    findEmpValueByNormalizedKey(emp, 'ifsccode'),
    findEmpValueByNormalizedKey(emp, 'ifsc_code'),
    findEmpValueByNormalizedKey(emp, 'ifsc')
  );

  const bankParts = [];
  if (accountNumber) bankParts.push(accountNumber);
  if (bankName) bankParts.push(bankName);
  if (branchName) bankParts.push(branchName);

  let combined = bankParts.join(', ');
  if (ifscCode) {
    combined = combined ? `${combined} (${ifscCode})` : ifscCode;
  }
  return combined;
}

export function ensureFormUTamilNaduMonthYearHeaderFields(formHeader) {
  if (!formHeader || typeof formHeader !== 'object') return formHeader;
  const fields = Array.isArray(formHeader.fields) ? [...formHeader.fields] : [];
  const hasMonth = fields.some(
    (f) =>
      f?.key === 'form_x_month' ||
      /^month\s*:?\s*$/i.test(String(f?.label || '').trim())
  );
  const hasYear = fields.some(
    (f) =>
      f?.key === 'form_x_year' ||
      /^year\s*:?\s*$/i.test(String(f?.label || '').trim())
  );
  if (!hasMonth) fields.push({ label: 'Month:', value: '', key: 'form_x_month' });
  if (!hasYear) fields.push({ label: 'Year:', value: '', key: 'form_x_year' });
  return { ...formHeader, fields };
}

function isMeaningfulExportCellValue(value) {
  const text = String(value ?? '').trim();
  if (!text) return false;
  if (/^enter\s+/i.test(text)) return false;
  if (/^select\s+/i.test(text)) return false;
  return true;
}

/**
 * Resolve a Form U TN row value for a template column label (export-safe).
 * Never use sourceHeaders[index] unless the alias bucket matches (avoids Photo→Email shifts).
 */
export function resolveFormUTamilNaduExportCellValue(row, templateHeader, sourceHeaders = [], headerIndex = null) {
  if (!row || typeof row !== 'object' || Array.isArray(row)) return '';
  const header = String(templateHeader || '').trim();
  if (!header) return '';

  if (Object.prototype.hasOwnProperty.call(row, header) && isMeaningfulExportCellValue(row[header])) {
    return row[header];
  }

  const targetBucket = formUTamilNaduHeaderAliasBucket(header);
  if (
    headerIndex != null &&
    headerIndex >= 0 &&
    sourceHeaders[headerIndex] &&
    formUTamilNaduHeaderAliasBucket(sourceHeaders[headerIndex]) === targetBucket
  ) {
    const sourceKey = sourceHeaders[headerIndex];
    if (Object.prototype.hasOwnProperty.call(row, sourceKey) && isMeaningfulExportCellValue(row[sourceKey])) {
      return row[sourceKey];
    }
  }

  const rowKeys = Object.keys(row).filter((k) => !String(k).startsWith('__'));

  for (const key of rowKeys) {
    if (formUTamilNaduHeaderAliasBucket(key) === targetBucket && isMeaningfulExportCellValue(row[key])) {
      return row[key];
    }
  }

  // Compatible aliases between modal FORMAT2 labels and TN template labels.
  if (targetBucket === 'workerid') {
    for (const key of rowKeys) {
      if (/worker identity|employee identification/i.test(key) && !/email/i.test(key)) {
        if (isMeaningfulExportCellValue(row[key])) return row[key];
      }
    }
  }
  if (targetBucket === 'aadhaar') {
    for (const key of rowKeys) {
      if (/aadhaar|aadhar/i.test(key) && isMeaningfulExportCellValue(row[key])) return row[key];
    }
  }
  if (targetBucket === 'doj') {
    for (const key of rowKeys) {
      if (
        /date on which joined|date of joining|date of entry|dateofjoining|\bdoj\b/i.test(key) &&
        isMeaningfulExportCellValue(row[key])
      ) {
        return row[key];
      }
    }
  }
  if (targetBucket === 'designation') {
    for (const key of rowKeys) {
      if (/designation|nature of work/i.test(key) && isMeaningfulExportCellValue(row[key])) {
        return row[key];
      }
    }
  }
  if (targetBucket === 'bank') {
    for (const key of rowKeys) {
      if (/bank/i.test(key) && !/address/i.test(key) && isMeaningfulExportCellValue(row[key])) {
        return row[key];
      }
    }
  }
  if (targetBucket === 'email') {
    const emailKeys = [
      'Email ID',
      'E Mail Id',
      'E Mail ID',
      'Email',
      'EmailID',
      'Email_ID',
      'email',
      'emailId',
      'email_id'
    ];
    for (const key of emailKeys) {
      if (Object.prototype.hasOwnProperty.call(row, key) && isMeaningfulExportCellValue(row[key])) {
        return row[key];
      }
    }
    for (const key of rowKeys) {
      if (/\bemail\b|e\s*mail|emailid|email_id/i.test(key) && isMeaningfulExportCellValue(row[key])) {
        return row[key];
      }
    }
  }
  if (targetBucket === 'mobile') {
    for (const key of rowKeys) {
      if (/mobile|phone|contact/i.test(key) && isMeaningfulExportCellValue(row[key])) return row[key];
    }
  }

  return '';
}
