import ExcelJS from 'exceljs';
import { parseEmployeeTabularSections } from './formXVIIIMPCombinedRegister';
import { ensureExcelJSDataRowsWithBorders } from '../../utils/excelTableBorders';
import { writeStatutoryHeaderFieldsToExcelJsWorksheet } from '../../utils/statutorySiteCompanyHeaders';

/** Gujarat Form A — FORMAT OF EMPLOYEE/ WORKMAN/ WORKER (Shops & Establishments). */

export const FORM_AGJ_GJ_CANONICAL_TABLE_HEADERS = [
  'Sr.No.',
  'Employees/ Workmen/ Worker Code',
  'Name',
  'Surname',
  'Gender',
  "Father's/ Spouse Name",
  'Date of Birth',
  'Nationality',
  'Education',
  'Date of Joining',
  'Designation',
  'Category Address',
  'Type of Employment',
  'Mobile',
  'UAN',
  'PAN',
  'ESIC IP',
  'AADHAAR',
  'Bank A/C No.',
  'Bank',
  'Branch (IFSC)',
  'Present Address',
  'Permanent Address',
];

export function formAGJGujaratHeaderNorm(txt) {
  return String(txt || '')
    .replace(/\r?\n/g, ' ')
    .replace(/['''`´]/g, '')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}

const stripLeadingNumber = (s) =>
  String(s || '')
    .replace(/^\d+[\.\)]\s*/, '')
    .trim();

export function formAGJHeaderAliasBucket(norm) {
  const n = String(norm || '').trim();
  if (!n) return '';
  if (/^sr\.?\s*no|^s\.?\s*no|serial|sl\.?\s*no|^no\.?$/.test(n)) return 'sno';
  // "Employee Code" (RJ Form A) and "Employees/ Workmen/ Worker Code" (GJ)
  if (
    /employee/.test(n) &&
    (/code|id/.test(n) || /workme|workman|worker|workmen/.test(n))
  ) {
    return 'employeeId';
  }
  if (n === 'name' || /^name$/.test(n)) return 'firstName';
  if (/surname/.test(n)) return 'surname';
  if (/gender|^sex$/.test(n)) return 'gender';
  if (/father/.test(n) || /spouse/.test(n) || /husband/.test(n)) return 'father';
  if (/date/.test(n) && (/birth|bi\b/.test(n) || /dob/.test(n))) return 'dob';
  if (/national/.test(n)) return 'nationality';
  if (/education/.test(n)) return 'education';
  if (/date/.test(n) && (/join|jo\b/.test(n) || /entry/.test(n))) return 'doj';
  if (/designat/.test(n)) return 'designation';
  if (/present/.test(n) && (/address|addr|\badd\b/.test(n))) return 'presentAddress';
  if (/permanent/.test(n) && (/address|addr|\badd\b/.test(n))) return 'permanentAddress';
  if (/aadha?ar/.test(n)) return 'aadhaar';
  if (
    /bank/.test(n) &&
    (/a\s*\/\s*c|a\s*c|account/.test(n) || (/ac/.test(n) && /no/.test(n)))
  ) {
    return 'bankAccount';
  }
  if (/branch/.test(n) || (/\bifsc\b/.test(n) && !/bank\s*a/.test(n))) return 'bankBranchIfsc';
  if (n === 'bank' || (/bank/.test(n) && !/account|ac|branch|ifsc|no/.test(n))) return 'bankName';
  // RJ Form A: Category (HS/S/SS/US)* — skill band, not address
  if (
    /category/.test(n) &&
    (/\(hs|hs\s*\/\s*s|\/ss\/|\/us\)|highly\s*skill|semi\s*skill|un\s*skill/.test(n) ||
      /\bhs\b/.test(n))
  ) {
    return 'skillCategory';
  }
  if (/category/.test(n) && (/address|addr/.test(n) || n.includes('category a'))) {
    return 'categoryAddress';
  }
  if (/type/.test(n) && (/employ|en\b|employment/.test(n))) return 'employmentType';
  if (/mobile|phone/.test(n)) return 'mobile';
  if (/^uan$/.test(n) || /\buan\b/.test(n)) return 'uan';
  if (/^pan$/.test(n) || /\bpan\b/.test(n)) return 'pan';
  if (/esic/.test(n)) return 'esic';
  if (/^lwf$/.test(n)) return 'lwf';
  return n;
}

export function headersIndicateFormAGJGujaratTable(tableHeaders) {
  if (!Array.isArray(tableHeaders) || tableHeaders.length < 6) return false;
  const joined = tableHeaders.map((h) => formAGJGujaratHeaderNorm(h)).join(' ');
  const buckets = tableHeaders.map((h) => formAGJHeaderAliasBucket(formAGJGujaratHeaderNorm(h)));
  const hasSno = buckets.includes('sno');
  const hasEducation = buckets.includes('education') || /education/.test(joined);
  const hasEmploymentType =
    buckets.includes('employmentType') || (/type/.test(joined) && /employ|en\b/.test(joined));
  const hasName = buckets.includes('firstName') || /\bname\b/.test(joined);
  return hasSno && hasEducation && hasEmploymentType && hasName;
}

export function isFormAGJGujaratContext(
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

  if (
    /maternity\s+benefit|register\s+of\s+muster\s+roll/.test(parts) &&
    !/gujarat|\b_gj\b|form_a_gj/.test(parts)
  ) {
    return false;
  }

  const hasFormB =
    /\bform[\s._-]*b\b/.test(parts) || /form[\s._-]*b[\s._-]*gj|form_b_gj/.test(parts);
  if (hasFormB && /gujarat|\b_gj\b/.test(parts)) return false;

  const hasFormC =
    /\bform[\s._-]*c\b/.test(parts) || /form[\s._-]*c[\s._-]*gj|form_c_gj/.test(parts);
  if (hasFormC && /gujarat|\b_gj\b/.test(parts)) return false;
  if (hasFormC && /register\s+of\s+deductions|deductions?\s+for\s+damage|damage\s+or\s+loss/.test(parts)) {
    return false;
  }

  const hasFormD =
    /\bform[\s._-]*d\b/.test(parts) || /form[\s._-]*d[\s._-]*gj|form_d_gj/.test(parts);
  if (hasFormD && /gujarat|\b_gj\b/.test(parts)) return false;
  if (hasFormD && /relay\s+or\s+set|summary\s+no\.?\s*of\s+days|register\s+keeper/.test(parts)) {
    return false;
  }

  const hasRajasthan = /rajasthan|\b_rj\b|form[\s._-]*a[\s._-]*rj|form_a_rj/.test(parts);
  if (hasRajasthan && /\bform[\s._-]*a\b/.test(parts)) return false;

  const hasGujarat = /gujarat|\b_gj\b|form[\s._-]*a[\s._-]*gj|form_a_gj/.test(parts);
  const hasFormA = /\bform[\s._-]*a\b/.test(parts);
  const hasEmployeeWorkmanFormat =
    /format\s+of\s+employee/.test(parts) ||
    (/employee\s*\/\s*workman\s*\/\s*worker/.test(parts) &&
      !/register\s+of\s+wages|rate\s+of\s+wage/.test(parts));

  if (hasGujarat && (hasFormA || hasEmployeeWorkmanFormat)) return true;

  if (Array.isArray(tableHeaders) && tableHeaders.length > 0) {
    const joined = tableHeaders.map((h) => formAGJGujaratHeaderNorm(h)).join('\n');
    if (
      hasGujarat &&
      /education\s+level/.test(joined) &&
      /type\s+of\s+employment/.test(joined)
    ) {
      return true;
    }
  }

  if (headersIndicateFormAGJGujaratTable(tableHeaders) && (hasGujarat || hasEmployeeWorkmanFormat)) {
    return true;
  }
  if (
    headersIndicateFormAGJGujaratTable(tableHeaders) &&
    hasFormA &&
    !/maternity|muster\s+roll|woman/.test(parts)
  ) {
    return true;
  }

  return false;
}

export function resolveFormAGJGujaratTableHeaders(tableHeaders) {
  const canonical = [...FORM_AGJ_GJ_CANONICAL_TABLE_HEADERS];
  const parsed = Array.isArray(tableHeaders)
    ? tableHeaders.map((h) => String(h || '').trim()).filter(Boolean)
    : [];
  if (!headersIndicateFormAGJGujaratTable(parsed)) {
    return canonical;
  }
  const merged = [...parsed];
  canonical.forEach((canonicalHeader) => {
    const bucket = formAGJHeaderAliasBucket(formAGJGujaratHeaderNorm(canonicalHeader));
    const hasBucket = merged.some(
      (h) => formAGJHeaderAliasBucket(formAGJGujaratHeaderNorm(h)) === bucket
    );
    if (!hasBucket) merged.push(canonicalHeader);
  });
  return merged;
}

export function isFormAGJEducationLevelHeader(h) {
  const s = formAGJGujaratHeaderNorm(stripLeadingNumber(h))
    .replace(/[^a-z0-9 ]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
  return /education/.test(s) && !/employ/.test(s);
}

export function isFormAGJTypeOfEmploymentHeader(h) {
  const s = formAGJGujaratHeaderNorm(stripLeadingNumber(h))
    .replace(/[^a-z0-9 ]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
  return /type/.test(s) && (/employ/.test(s) || /\ben\b/.test(s) || /employment/.test(s));
}

function unwrapEmployeeRecord(emp) {
  return emp?.Employee || emp?.employee || emp;
}

function pickEmployeeValue(emp, keys) {
  const src = unwrapEmployeeRecord(emp);
  if (!src || typeof src !== 'object') return '';
  for (let i = 0; i < keys.length; i += 1) {
    const raw = src[keys[i]];
    if (raw == null || raw === '') continue;
    if (typeof raw === 'object') {
      const nested = String(raw.displayValue ?? raw.name ?? raw.Name ?? '').trim();
      if (nested) return nested;
      continue;
    }
    const text = String(raw).trim();
    if (text) return text;
  }
  return '';
}

/** Format Education Details from Zoho People `tabularSections`. */
export function formatEducationLevelFromEmployee(emp) {
  const tabs = parseEmployeeTabularSections(emp);
  if (!tabs || typeof tabs !== 'object') return '';

  const eduKey =
    Object.keys(tabs).find((k) => /education\s*details/i.test(String(k || ''))) || 'Education Details';
  const eduList = tabs[eduKey];
  if (!Array.isArray(eduList)) return '';

  const parts = [];
  eduList.forEach((edu) => {
    if (!edu || typeof edu !== 'object') return;
    const degree = String(
      edu.Degree ?? edu.degree ?? edu.Qualification ?? edu.qualification ?? ''
    ).trim();
    const spec = String(
      edu.Specialization ??
        edu.specialization ??
        edu.Field ??
        edu.field ??
        edu['Field of Study'] ??
        ''
    ).trim();
    const inst = String(
      edu.University ??
        edu.university ??
        edu.College ??
        edu.college ??
        edu.Institution ??
        edu.institution ??
        ''
    ).trim();
    const line = [degree, spec].filter(Boolean).join(' - ');
    const full = [line, inst].filter(Boolean).join(', ');
    if (full) parts.push(full);
  });

  return parts.join('; ');
}

/** Read Type of Employment from Zoho People `Employee_type`. */
export function readEmployeeTypeFromEmployee(emp) {
  return pickEmployeeValue(emp, [
    'Employee_type',
    'Employee Type',
    'EmployeeType',
    'employee_type',
    'employeeType',
    'Employee_type.displayValue',
    'Employee Type.displayValue',
  ]);
}

function readEmployeeId(emp) {
  return pickEmployeeValue(emp, [
    'EmployeeID',
    'Employee ID',
    'Employee_ID',
    'EmployeeId',
    'employeeId',
    'Employee.ID',
    'erecno',
    'Erecno',
    'Zoho_ID',
  ]);
}

function readMiddleName(emp) {
  return pickEmployeeValue(emp, [
    'MiddleName',
    'Middle Name',
    'Middle_Name',
    'middleName',
    'middle_name',
  ]);
}

function readFirstName(emp) {
  const first = pickEmployeeValue(emp, [
    'FirstName',
    'First Name',
    'First_Name',
    'firstName',
    'first_name',
  ]);
  const middle = readMiddleName(emp);
  const combined = [first, middle].filter(Boolean).join(' ').trim();
  if (combined) return combined;
  return pickEmployeeValue(emp, ['Name', 'Name1', 'Employee Name', 'Employee_Name', 'DisplayName']);
}

function readSurname(emp) {
  return pickEmployeeValue(emp, ['LastName', 'Last Name', 'Last_Name', 'lastName', 'Surname', 'surname']);
}

function readGender(emp) {
  return pickEmployeeValue(emp, ['Sex', 'Gender', 'sex', 'gender']);
}

function readFatherName(emp) {
  return pickEmployeeValue(emp, [
    'Father_s_Name',
    'Father_s Name',
    'FatherName',
    'Father Name',
    "Father's Name",
    'SpouseName',
    'Spouse Name',
    'HusbandName',
    'Husband Name',
  ]);
}

function readDob(emp) {
  return pickEmployeeValue(emp, [
    'Date_of_birth',
    'Date of Birth',
    'DateofBirth',
    'Dateofbirth',
    'DOB',
    'dob',
  ]);
}

function readNationality(emp) {
  return pickEmployeeValue(emp, ['Nationality', 'nationality', 'Country', 'Country1']);
}

function readDoj(emp) {
  return pickEmployeeValue(emp, [
    'Dateofjoining',
    'DateofJoining',
    'Date of Joining',
    'Date_of_Joining',
    'DateofentryintoService',
    'Date of entry into service',
  ]);
}

function readDesignation(emp) {
  return pickEmployeeValue(emp, [
    'Designation',
    'designation',
    'Designation.displayValue',
    'Department',
    'department',
  ]);
}

function readCategoryAddress(emp) {
  const src = unwrapEmployeeRecord(emp);
  const city = pickEmployeeValue(src, ['City', 'City1', 'City2', 'Present City', 'Work_location', 'LocationName']);
  const state = pickEmployeeValue(src, ['State', 'State1']);
  const address = pickEmployeeValue(src, [
    'Present_Address',
    'Present Address',
    'PresentAddress',
    'Address_Line_1',
    'Address Line 1',
  ]);
  return [city, state].filter(Boolean).join(', ') || address;
}

/** Normalize skill band to Form A Category (HS/S/SS/US)*. */
export function normalizeFormASkillCategory(raw) {
  const t = String(raw || '')
    .replace(/\r?\n/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
  if (!t) return '';
  const compact = t.toLowerCase().replace(/[^a-z]/g, '');
  if (compact === 'hs' || compact === 'highlyskilled' || compact.startsWith('highlyskill')) return 'HS';
  if (compact === 'ss' || compact === 'semiskilled' || compact.startsWith('semiskill')) return 'SS';
  if (compact === 'us' || compact === 'unskilled' || compact.startsWith('unskill')) return 'US';
  if (compact === 's' || compact === 'skilled' || (compact.startsWith('skill') && !compact.includes('semi') && !compact.includes('un') && !compact.includes('high'))) {
    return 'S';
  }
  if (/^(HS|S|SS|US)$/i.test(t)) return t.toUpperCase();
  return t;
}

function readSkillCategory(emp) {
  const raw = pickEmployeeValue(emp, [
    'Skill_Category',
    'Skill Category',
    'SkillCategory',
    'Skill_Level',
    'Skill Level',
    'SkillLevel',
    'Worker_Category',
    'Worker Category',
    'WorkerCategory',
    'Emp_Category',
    'Employee Category',
    'EmployeeCategory',
    'Category_HS_S_SS_US',
    'Skill',
  ]);
  return normalizeFormASkillCategory(raw);
}

function readMobile(emp) {
  return pickEmployeeValue(emp, [
    'Mobile',
    'Mobile_Number',
    'Mobile Number',
    'Phone',
    'Phone_Number',
    'Phone Number',
  ]);
}

function readUan(emp) {
  return pickEmployeeValue(emp, ['UAN_Number', 'UAN Number', 'UAN', 'uan']);
}

function readPan(emp) {
  return pickEmployeeValue(emp, ['PAN', 'PAN_Number', 'PAN Number', 'Pan', 'pan']);
}

function readEsic(emp) {
  return pickEmployeeValue(emp, [
    "Employee's State Insurance Corporation No.",
    'ESIC_Number',
    'ESIC Number',
    'ESICNo',
    'ESIC_No',
    'ESI_Number',
    'ESI Number',
  ]);
}

function readAadhaar(emp) {
  return pickEmployeeValue(emp, [
    'Aadhaar_Number',
    'Aadhaar Number',
    'AadhaarNo',
    'Aadhaar No.',
    'AadhaarNo',
    'aadhaarNo',
    'AADHAAR',
    'Aadhaar',
  ]);
}

function readBankAccount(emp) {
  return pickEmployeeValue(emp, [
    'Account_Number',
    'Account Number',
    'AccountNumber',
    'accountNumber',
    'Bank_Account_Number',
    'Bank Account Number',
  ]);
}

function readBankName(emp) {
  const src = unwrapEmployeeRecord(emp);
  const direct = pickEmployeeValue(src, [
    'Bank_Name',
    'Bank Name',
    'BankName',
    'bankName',
    'BANK_NAME',
  ]);
  if (direct) return direct;
  const bankFields = Object.keys(src || {}).filter((key) => {
    const keyLower = String(key || '').toLowerCase();
    return (
      keyLower.includes('bank') &&
      !keyLower.includes('account') &&
      !keyLower.includes('branch') &&
      !keyLower.includes('ifsc') &&
      !keyLower.includes('code')
    );
  });
  for (let i = 0; i < bankFields.length; i += 1) {
    const value = String(src[bankFields[i]] ?? '').trim();
    if (value) return value;
  }
  return '';
}

function readBankBranchIfsc(emp) {
  const branch = pickEmployeeValue(emp, [
    'Branch_Name',
    'Branch Name',
    'BranchName',
    'branchName',
    'bankBranch',
    'Bank Branch',
    'Bank_Branch',
  ]);
  const ifsc = pickEmployeeValue(emp, [
    'IFSC_Code',
    'IFSC Code',
    'IFSCCode',
    'ifscCode',
    'IFSC',
    'ifsc',
  ]);
  if (branch && ifsc) return `${branch} (${ifsc})`;
  return branch || ifsc || '';
}

function buildAddressFromParts(parts) {
  return parts.map((value) => String(value || '').trim()).filter(Boolean).join(', ');
}

function readPresentAddress(emp) {
  const src = unwrapEmployeeRecord(emp);
  const direct = pickEmployeeValue(src, [
    'Present_Address',
    'Present Address',
    'PresentAddress',
    'presentAddress',
    'presentAddressLine1',
    'Present Address Line 1',
  ]);
  if (direct) return direct;

  const line1 = pickEmployeeValue(src, [
    'Address_Line_1',
    'Address Line 1',
    'AddressLine1',
    'presentAddressLine1',
    'Present_Address_Line_1',
  ]);
  const line2 = pickEmployeeValue(src, [
    'Address_Line_2',
    'Address Line 2',
    'AddressLine2',
    'presentAddressLine2',
  ]);
  const city = pickEmployeeValue(src, [
    'City1',
    'City 1',
    'City',
    'presentCity',
    'Present City',
  ]);
  const state = pickEmployeeValue(src, ['State', 'State1', 'presentState', 'Present State']);
  const postal = pickEmployeeValue(src, ['Pincode', 'PostalCode', 'presentPostalCode', 'Zip']);
  const country = pickEmployeeValue(src, ['Country', 'Country1', 'presentCountry']);
  return (
    buildAddressFromParts([line1, line2, city, state, postal, country]) ||
    pickEmployeeValue(src, ['Address', 'address']) ||
    ''
  );
}

function readPermanentAddress(emp) {
  const src = unwrapEmployeeRecord(emp);
  const direct = pickEmployeeValue(src, [
    'Permanent_Address',
    'Permanent Address',
    'Permanentaddress',
    'PermanentAddress',
    'permanentAddress',
    'permanentAddressLine1',
  ]);
  if (direct) return direct;

  const line1 = pickEmployeeValue(src, [
    'Permanent_Address_Line_1',
    'Address_Line_1',
    'Address Line 1',
    'AddressLine1',
    'permanentAddressLine1',
  ]);
  const line2 = pickEmployeeValue(src, [
    'Permanent_Address_Line_2',
    'Address_Line_2',
    'Address Line 2',
    'AddressLine2',
    'permanentAddressLine2',
  ]);
  const city = pickEmployeeValue(src, [
    'City2',
    'City 2',
    'City',
    'permanentCity',
    'Permanent City',
  ]);
  const state = pickEmployeeValue(src, ['State1', 'State', 'permanentState', 'Permanent State']);
  const postal = pickEmployeeValue(src, ['Pincode1', 'permanentPostalCode', 'PostalCode']);
  const country = pickEmployeeValue(src, ['Country1', 'Country', 'permanentCountry']);
  const built = buildAddressFromParts([line1, line2, city, state, postal, country]);
  return built || readPresentAddress(src);
}

function bucketValueForEmployee(bucket, emp, helpers = {}) {
  const { formatStatutoryDateDisplay = (v) => String(v || '').trim(), rowIndex = 0 } = helpers;
  switch (bucket) {
    case 'sno':
      return String(rowIndex + 1);
    case 'employeeId':
      return readEmployeeId(emp);
    case 'firstName':
      return readFirstName(emp);
    case 'surname':
      return readSurname(emp);
    case 'gender':
      return readGender(emp);
    case 'father':
      return readFatherName(emp);
    case 'dob': {
      const raw = readDob(emp);
      return raw ? formatStatutoryDateDisplay(raw) : '';
    }
    case 'nationality':
      return readNationality(emp);
    case 'education':
      return formatEducationLevelFromEmployee(emp);
    case 'doj': {
      const raw = readDoj(emp);
      return raw ? formatStatutoryDateDisplay(raw) : '';
    }
    case 'designation':
      return readDesignation(emp);
    case 'categoryAddress':
      return readCategoryAddress(emp);
    case 'skillCategory':
      return readSkillCategory(emp);
    case 'employmentType':
      return readEmployeeTypeFromEmployee(emp);
    case 'mobile':
      return readMobile(emp);
    case 'uan':
      return readUan(emp);
    case 'pan':
      return readPan(emp);
    case 'esic':
      return readEsic(emp);
    case 'aadhaar':
      return readAadhaar(emp);
    case 'bankAccount':
      return readBankAccount(emp);
    case 'bankName':
      return readBankName(emp);
    case 'bankBranchIfsc':
      return readBankBranchIfsc(emp);
    case 'presentAddress':
      return readPresentAddress(emp);
    case 'permanentAddress':
      return readPermanentAddress(emp);
    default:
      return '';
  }
}

export function getFormAGJGujaratRowValueForHeader(row, header, rowIndex = 0) {
  if (!row || !header) return '';
  const direct = row[header];
  if (direct != null && String(direct).trim() !== '') return String(direct).trim();
  const bucket = formAGJHeaderAliasBucket(formAGJGujaratHeaderNorm(header));
  for (const [k, v] of Object.entries(row)) {
    if (
      formAGJHeaderAliasBucket(formAGJGujaratHeaderNorm(k)) === bucket &&
      v != null &&
      String(v).trim() !== ''
    ) {
      return String(v).trim();
    }
  }
  if (bucket === 'sno') return String(rowIndex + 1);
  return '';
}

export function rowHasMeaningfulFormAGJGujaratExportData(row, headers) {
  const hdrs = resolveFormAGJGujaratTableHeaders(headers);
  return hdrs.some((header) => {
    const bucket = formAGJHeaderAliasBucket(formAGJGujaratHeaderNorm(header));
    if (bucket === 'sno') return false;
    return getFormAGJGujaratRowValueForHeader(row, header) !== '';
  });
}

export function mapFormAGJGujaratRowsFromEmployees(employees, headers, helpers = {}) {
  const hdrs = resolveFormAGJGujaratTableHeaders(headers);
  const list = Array.isArray(employees) ? employees : [];
  const {
    sanitizeValue = (v) => String(v ?? '').trim(),
    formatStatutoryDateDisplay = (v) => String(v || '').trim(),
    rowIndexOffset = 0,
  } = helpers;

  return list.map((empItem, rowIndex) => {
    const emp = unwrapEmployeeRecord(empItem);
    const globalRowIndex = rowIndexOffset + rowIndex;
    const row = {};
    hdrs.forEach((header) => {
      const bucket = formAGJHeaderAliasBucket(formAGJGujaratHeaderNorm(header));
      row[header] = sanitizeValue(
        bucketValueForEmployee(bucket, emp, {
          formatStatutoryDateDisplay,
          rowIndex: globalRowIndex,
        })
      );
    });
    const name = [readFirstName(emp), readSurname(emp)].filter(Boolean).join(' ').trim();
    if (name) row.__employeeLookupName = name;
    const id = readEmployeeId(emp);
    if (id) row.__employeeLookupId = id;
    return row;
  });
}

export function applyFormAGJGujaratEmployeeToRow(row, emp, headers, helpers = {}) {
  const mapped = mapFormAGJGujaratRowsFromEmployees([emp], headers, {
    ...helpers,
    rowIndexOffset: helpers.rowIndex ?? 0,
  })[0];
  const out = { ...row };
  Object.keys(mapped || {}).forEach((key) => {
    if (String(key).startsWith('__')) {
      out[key] = mapped[key];
      return;
    }
    const val = mapped[key];
    if (val != null && String(val).trim() !== '') {
      out[key] = val;
    }
  });
  return out;
}

export function remapFormAGJGujaratRowsToHeaders(rows, sourceHeaders, targetHeaders) {
  const src = Array.isArray(sourceHeaders) ? sourceHeaders : [];
  const tgt = resolveFormAGJGujaratTableHeaders(targetHeaders);
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

const excelCellValueToString = (val) => {
  if (val == null) return '';
  if (typeof val === 'string' || typeof val === 'number' || typeof val === 'boolean') return String(val);
  if (val instanceof Date) return val.toISOString();
  if (typeof val === 'object') {
    if (Array.isArray(val.richText)) return val.richText.map((rt) => rt?.text || '').join('');
    if (val.text != null) return String(val.text);
    if (val.result != null) return String(val.result);
  }
  return '';
};

function excelCellLooksLikeSerialHeader(text) {
  const t = formAGJGujaratHeaderNorm(text).replace(/[^a-z0-9 ]/g, ' ').replace(/\s+/g, ' ').trim();
  return /^sr\.?\s*no|^s\.?\s*no|serial|sl\.?\s*no/.test(t);
}

function buildMergeTopLeftResolver(worksheet) {
  const cache = new Map();
  return (r, c) => {
    const key = `${r}:${c}`;
    if (cache.has(key)) return cache.get(key);
    let topLeft = { r, c };
    listFormAGJWorksheetMerges(worksheet).forEach((range) => {
      const parts = String(range || '').split(':');
      if (parts.length !== 2) return;
      const start = parseFormAGJMergeRef(parts[0]);
      const end = parseFormAGJMergeRef(parts[1]);
      if (!start || !end) return;
      const r1 = Math.min(start.row, end.row);
      const r2 = Math.max(start.row, end.row);
      const c1 = Math.min(start.col, end.col);
      const c2 = Math.max(start.col, end.col);
      if (r >= r1 && r <= r2 && c >= c1 && c <= c2) {
        topLeft = { r: r1, c: c1 };
      }
    });
    cache.set(key, topLeft);
    return topLeft;
  };
}

function detectFormAGJGujaratTableLayout(worksheet, hints = {}) {
  const getMergeTopLeft = buildMergeTopLeftResolver(worksheet);
  const getMergedAwareCellText = (r, c) => {
    const tl = getMergeTopLeft(r, c);
    return excelCellValueToString(worksheet.getCell(tl.r, tl.c)?.value).trim();
  };

  let headerRow =
    hints.parsedHeaderRowIndex != null && hints.parsedHeaderRowIndex >= 0
      ? hints.parsedHeaderRowIndex + 1
      : -1;
  let startCol =
    hints.parsedTableStartCol != null && hints.parsedTableStartCol >= 0
      ? hints.parsedTableStartCol + 1
      : 1;

  if (headerRow < 1) {
    const maxScanRows = Math.max(35, worksheet.rowCount + 5);
    for (let r = 1; r <= maxScanRows; r += 1) {
      for (let c = 1; c <= 25; c += 1) {
        if (excelCellLooksLikeSerialHeader(getMergedAwareCellText(r, c))) {
          headerRow = r;
          startCol = c;
          break;
        }
      }
      if (headerRow > 0) break;
    }
  }
  if (headerRow < 1) return null;

  const seenBuckets = new Set();
  const templateCols = [];
  const maxTemplateCols = 45;
  const scanColEnd = Math.max(startCol + 50, 55);
  for (let c = startCol; c < scanColEnd; c += 1) {
    const label = getMergedAwareCellText(headerRow, c);
    if (!label || /^\d+$/.test(label)) continue;
    const bucket = formAGJHeaderAliasBucket(formAGJGujaratHeaderNorm(label));
    if (!bucket || seenBuckets.has(bucket)) continue;
    seenBuckets.add(bucket);
    templateCols.push({ col: c, label, bucket });
    if (templateCols.length >= maxTemplateCols) break;
  }
  if (templateCols.length < 4) return null;

  let dataStartRow =
    hints.parsedDataStartIndex != null && hints.parsedDataStartIndex >= 0
      ? hints.parsedDataStartIndex + 1
      : headerRow + 1;

  let seqHits = 0;
  for (let i = 0; i < templateCols.length; i += 1) {
    const t = getMergedAwareCellText(dataStartRow, templateCols[i].col).replace(/\s+/g, '').trim();
    if (t === String(i + 1)) seqHits += 1;
  }
  if (seqHits >= Math.max(3, Math.floor(templateCols.length * 0.35))) {
    dataStartRow = headerRow + 2;
  }

  return { headerRow, dataStartRow, templateCols };
}

/** Title band: centered merge E–Z, rows 1–3 (matches Gujarat Form A template). */
const FORM_AGJ_TITLE_MERGE_COL_FROM = 5; // E
const FORM_AGJ_TITLE_MERGE_COL_TO = 26; // Z
const FORM_AGJ_SCHEDULE_ROW = 1;
const FORM_AGJ_TITLE_MAIN_ROW = 2;
const FORM_AGJ_TITLE_SUB_ROW = 3;
const FORM_AGJ_DEFAULT_SCHEDULE = 'SCHEDULE';
const FORM_AGJ_DEFAULT_MAIN_TITLE = 'FORM A';
const FORM_AGJ_DEFAULT_SUBTITLE = 'FORMAT OF EMPLOYEE/ WORKMAN/ WORKER';
const FORM_AGJ_DEFAULT_SEE_RULE = '[See rule 2(1)]';

function normalizeFormAGJSubtitleSingleLine(text) {
  const cleaned = String(text || FORM_AGJ_DEFAULT_SUBTITLE)
    .replace(/\r?\n/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
  if (isPlaceholderBandText(cleaned)) return FORM_AGJ_DEFAULT_SUBTITLE;
  return cleaned || FORM_AGJ_DEFAULT_SUBTITLE;
}

function isPlaceholderBandText(text) {
  const raw = String(text || '').trim();
  if (!raw) return true;
  if (/^[-–—_.\s]+$/.test(raw)) return true;
  if (/^[-–—]{3,}$/.test(raw.replace(/\s/g, ''))) return true;
  return false;
}

function sheetTextLooksLikeFormAGJMainTitle(text) {
  const n = formAGJGujaratHeaderNorm(text).replace(/[^a-z0-9 ]/g, ' ').replace(/\s+/g, ' ').trim();
  return n === 'form a' || /^form\s*a$/.test(n);
}

function sheetTextLooksLikeFormAGJSubtitle(text) {
  if (isPlaceholderBandText(text)) return false;
  const n = formAGJGujaratHeaderNorm(text);
  return /format\s+of\s+employee/.test(n) && /workman|worker/.test(n);
}

function sheetTextLooksLikeFormAGJSchedule(text) {
  if (isPlaceholderBandText(text)) return false;
  const n = formAGJGujaratHeaderNorm(text).replace(/[^a-z0-9 ]/g, ' ').replace(/\s+/g, ' ').trim();
  return n === 'schedule' || /^schedule(\s+[ivxlc\d]+)?$/.test(n) || n === 'sched';
}

function sheetTextLooksLikeFormAGJSeeRule(text) {
  if (isPlaceholderBandText(text)) return false;
  const n = formAGJGujaratHeaderNorm(text);
  return /see\s*rule\s*2/.test(n);
}

function sheetTextLooksLikeFormAGJPartA(text) {
  if (isPlaceholderBandText(text)) return false;
  const n = formAGJGujaratHeaderNorm(text);
  return /part\s*[-\s]?a\b/.test(n) && /establish/.test(n);
}

function formAGJColLettersToNumber(letters) {
  let n = 0;
  const s = String(letters || '').toUpperCase();
  for (let i = 0; i < s.length; i += 1) {
    n = n * 26 + (s.charCodeAt(i) - 64);
  }
  return n;
}

function parseFormAGJMergeRef(ref) {
  const m = String(ref || '')
    .replace(/\$/g, '')
    .match(/^([A-Z]+)(\d+)$/i);
  if (!m) return null;
  return { col: formAGJColLettersToNumber(m[1]), row: parseInt(m[2], 10) };
}

/** Collect merge ranges from model.merges and internal _merges. */
function listFormAGJWorksheetMerges(worksheet) {
  const out = [];
  const seen = new Set();
  const pushRange = (range) => {
    const key = String(range || '').replace(/\$/g, '').toUpperCase();
    if (!key || seen.has(key)) return;
    seen.add(key);
    out.push(key);
  };
  const modelMerges = worksheet?.model?.merges;
  if (Array.isArray(modelMerges)) {
    modelMerges.forEach((range) => pushRange(range));
  }
  const internal = worksheet?._merges;
  if (internal && typeof internal === 'object') {
    Object.keys(internal).forEach((addr) => {
      const dim = internal[addr];
      if (dim && dim.range) pushRange(dim.range);
      else if (dim && dim.top != null && dim.left != null && dim.bottom != null && dim.right != null) {
        const a = worksheet.getCell(dim.top, dim.left)?.address;
        const b = worksheet.getCell(dim.bottom, dim.right)?.address;
        if (a && b) pushRange(`${a}:${b}`);
      }
    });
  }
  return out;
}

/**
 * Unmerge any ranges overlapping the title band.
 * Must cover merges wider than S (e.g. E2:Z2) — otherwise writes to S land on the
 * merge master in E and stay visually centered over E–F.
 */
function unmergeFormAGJTitleBandMerges(worksheet, rowFrom, rowTo, colFrom, colTo) {
  if (!worksheet) return;
  for (let pass = 0; pass < 16; pass += 1) {
    let removed = 0;
    listFormAGJWorksheetMerges(worksheet).forEach((range) => {
      const parts = String(range || '').split(':');
      if (parts.length !== 2) return;
      const start = parseFormAGJMergeRef(parts[0]);
      const end = parseFormAGJMergeRef(parts[1]);
      if (!start || !end) return;
      const r1 = Math.min(start.row, end.row);
      const r2 = Math.max(start.row, end.row);
      const c1 = Math.min(start.col, end.col);
      const c2 = Math.max(start.col, end.col);
      if (r2 < rowFrom || r1 > rowTo || c2 < colFrom || c1 > colTo) return;
      try {
        worksheet.unMergeCells(range);
        removed += 1;
      } catch (_) {
        try {
          worksheet.unMergeCells(r1, c1, r2, c2);
          removed += 1;
        } catch (_2) {
          /* ignore */
        }
      }
    });
    if (!removed) break;
  }
}

function sheetTextLooksLikeFormAGJTitleBand(text) {
  return (
    sheetTextLooksLikeFormAGJMainTitle(text) ||
    sheetTextLooksLikeFormAGJSubtitle(text) ||
    sheetTextLooksLikeFormAGJSchedule(text)
  );
}

/** Title-band lines omitted from downloaded Excel/PDF (establishment fields stay). */
export function isFormAGJGujaratExportTitleBandText(text) {
  if (isPlaceholderBandText(text)) return false;
  if (sheetTextLooksLikeFormAGJSeeRule(text)) return true;
  if (sheetTextLooksLikeFormAGJTitleBand(text)) return true;
  const n = formAGJGujaratHeaderNorm(text);
  if (/see\s*rule\s*2/.test(n) && /schedule/.test(n)) return true;
  return false;
}

function buildFormAGJTitleBandClearPredicate() {
  const cellLooksLikeTitleToClear = (raw) => {
    if (!raw) return false;
    if (sheetTextLooksLikeFormAGJSeeRule(raw) || sheetTextLooksLikeFormAGJPartA(raw)) return false;
    if (sheetTextLooksLikeFormAGJTitleBand(raw) || isPlaceholderBandText(raw)) return true;
    const n = formAGJGujaratHeaderNorm(raw);
    if (/^schedule(\s+[ivxlc\d]+)?$/.test(n.replace(/[^a-z0-9 ]/g, ' ').replace(/\s+/g, ' ').trim())) {
      return true;
    }
    if (/^form\s*a\b/.test(n) && n.length < 20) return true;
    if (/format\s+of\s+employee/.test(n)) return true;
    return false;
  };
  return cellLooksLikeTitleToClear;
}

/** Remove SCHEDULE / FORM A / FORMAT title rows from the export worksheet. */
export function clearFormAGJGujaratExportTitleBands(worksheet) {
  if (!worksheet) return;
  const getMergeTopLeft = buildMergeTopLeftResolver(worksheet);
  const getMergedAwareCellText = (r, c) => {
    const tl = getMergeTopLeft(r, c);
    return excelCellValueToString(worksheet.getCell(tl.r, tl.c)?.value).trim();
  };
  const scanRowFrom = 1;
  const scanRowTo = 15;
  const unmergeColTo = Math.max(FORM_AGJ_TITLE_MERGE_COL_TO + 14, 40);
  const cellLooksLikeTitleToClear = buildFormAGJTitleBandClearPredicate();
  const clearedMasters = new Set();

  unmergeFormAGJTitleBandMerges(worksheet, scanRowFrom, scanRowTo, 1, unmergeColTo);

  const shouldClearCellText = (raw) =>
    cellLooksLikeTitleToClear(raw) || sheetTextLooksLikeFormAGJSeeRule(raw);

  const clearTitleAt = (row, col) => {
    const raw = getMergedAwareCellText(row, col);
    if (!shouldClearCellText(raw)) return;
    const tl = getMergeTopLeft(row, col);
    const key = `${tl.r}:${tl.c}`;
    if (clearedMasters.has(key)) return;
    clearedMasters.add(key);
    unmergeFormAGJTitleBandMerges(worksheet, tl.r, tl.r, 1, unmergeColTo);
    worksheet.getCell(tl.r, tl.c).value = null;
  };

  for (let r = scanRowFrom; r <= scanRowTo; r += 1) {
    for (let c = 1; c <= unmergeColTo; c += 1) {
      clearTitleAt(r, c);
    }
  }
}

/** Last pass on any download workbook that looks like Gujarat Form A. */
export function finalizeFormAGJGujaratExportTitleBandsOnWorksheet(worksheet) {
  if (!worksheet) return;
  if (!detectFormAGJGujaratTableLayout(worksheet, {})) return;
  clearFormAGJGujaratExportTitleBands(worksheet);
}

/**
 * Form A GJ title band (rows 1–3, centered E–Z):
 * SCHEDULE → FORM A → FORMAT OF EMPLOYEE/ WORKMAN/ WORKER.
 */
export function writeFormAGJGujaratTitleBandsInColumnS(worksheet, parsedFormHeader = {}) {
  if (!worksheet) return;
  const getMergeTopLeft = buildMergeTopLeftResolver(worksheet);
  const getMergedAwareCellText = (r, c) => {
    const tl = getMergeTopLeft(r, c);
    return excelCellValueToString(worksheet.getCell(tl.r, tl.c)?.value).trim();
  };

  const parsedTitle = String(parsedFormHeader?.title || '').trim();
  const parsedSubtitle = String(parsedFormHeader?.subtitle || '').trim();
  const scheduleText = FORM_AGJ_DEFAULT_SCHEDULE;
  const mainText =
    parsedTitle && !isPlaceholderBandText(parsedTitle) && sheetTextLooksLikeFormAGJMainTitle(parsedTitle)
      ? parsedTitle
      : FORM_AGJ_DEFAULT_MAIN_TITLE;
  const subText = normalizeFormAGJSubtitleSingleLine(
    parsedSubtitle && sheetTextLooksLikeFormAGJSubtitle(parsedSubtitle)
      ? parsedSubtitle
      : FORM_AGJ_DEFAULT_SUBTITLE
  );

  let seeRuleText = FORM_AGJ_DEFAULT_SEE_RULE;
  for (let r = 1; r <= 8; r += 1) {
    for (let c = 1; c <= 40; c += 1) {
      const raw = getMergedAwareCellText(r, c);
      if (raw && sheetTextLooksLikeFormAGJSeeRule(raw)) {
        seeRuleText = raw;
        break;
      }
    }
  }

  const scheduleRow = FORM_AGJ_SCHEDULE_ROW;
  const mainRow = FORM_AGJ_TITLE_MAIN_ROW;
  const subRow = FORM_AGJ_TITLE_SUB_ROW;
  const scanRowFrom = 1;
  const scanRowTo = 8;
  const unmergeColTo = Math.max(FORM_AGJ_TITLE_MERGE_COL_TO + 14, 40);

  unmergeFormAGJTitleBandMerges(worksheet, scanRowFrom, scanRowTo, 1, unmergeColTo);

  const cellLooksLikeTitleToClear = buildFormAGJTitleBandClearPredicate();

  const clearTitleDuplicatesInRow = (row) => {
    for (let c = 1; c <= unmergeColTo; c += 1) {
      const tl = getMergeTopLeft(row, c);
      if (tl.r !== row || tl.c !== c) continue;
      const raw = excelCellValueToString(worksheet.getCell(row, c)?.value).trim();
      if (!cellLooksLikeTitleToClear(raw)) continue;
      worksheet.getCell(row, c).value = null;
    }
  };

  for (let r = scanRowFrom; r <= scanRowTo; r += 1) {
    clearTitleDuplicatesInRow(r);
  }

  const writeCenteredBand = (row, text, kind) => {
    if (!row || row < 1 || !text) return;
    unmergeFormAGJTitleBandMerges(worksheet, row, row, 1, unmergeColTo);
    clearTitleDuplicatesInRow(row);
    const colFrom = FORM_AGJ_TITLE_MERGE_COL_FROM;
    const colTo = FORM_AGJ_TITLE_MERGE_COL_TO;
    try {
      worksheet.mergeCells(row, colFrom, row, colTo);
    } catch (_) {
      /* ignore */
    }
    const cell = worksheet.getCell(row, colFrom);
    cell.value = text;
    const isSubtitle = kind === 'subtitle';
    cell.alignment = {
      horizontal: 'center',
      vertical: 'middle',
      wrapText: isSubtitle,
      shrinkToFit: false,
    };
    cell.font = {
      ...(cell.font || {}),
      bold: true,
      size: isSubtitle ? 9 : 11,
    };
    const wsRow = worksheet.getRow(row);
    if (wsRow) wsRow.height = isSubtitle ? 24 : 18;
  };

  writeCenteredBand(scheduleRow, scheduleText, 'schedule');
  writeCenteredBand(mainRow, mainText, 'main');
  writeCenteredBand(subRow, subText, 'subtitle');

  if (seeRuleText) {
    try {
      worksheet.unMergeCells(scheduleRow, 6, scheduleRow, 10);
    } catch (_) {
      /* ignore */
    }
    const ruleCell = worksheet.getCell(scheduleRow, 6);
    ruleCell.value = seeRuleText;
    ruleCell.alignment = { horizontal: 'left', vertical: 'middle', wrapText: true };
    ruleCell.font = { ...(ruleCell.font || {}), bold: false, size: 9 };
  }

  // Remove stale copies that lived on rows 4–5 in older templates.
  for (let r = subRow + 1; r <= subRow + 3; r += 1) {
    let onlyTitle = true;
    for (let c = 1; c <= unmergeColTo; c += 1) {
      const raw = getMergedAwareCellText(r, c);
      if (!raw) continue;
      if (sheetTextLooksLikeFormAGJPartA(raw)) {
        onlyTitle = false;
        break;
      }
      if (!cellLooksLikeTitleToClear(raw)) {
        onlyTitle = false;
        break;
      }
    }
    if (onlyTitle) clearTitleDuplicatesInRow(r);
  }
}

export function filterFormAGJGujaratExportRows(rows, headers) {
  const hdrs = resolveFormAGJGujaratTableHeaders(headers);
  return (Array.isArray(rows) ? rows : []).filter((row) =>
    rowHasMeaningfulFormAGJGujaratExportData(row, hdrs)
  );
}

export function enrichFormAGJGujaratRowsForExport(mappedData, headers, employeesOverride = null, helpers = {}) {
  const hdrs = resolveFormAGJGujaratTableHeaders(headers);
  const rows = filterFormAGJGujaratExportRows(
    remapFormAGJGujaratRowsToHeaders(Array.isArray(mappedData) ? mappedData : [], hdrs, hdrs),
    hdrs
  );
  const employees = Array.isArray(employeesOverride) ? employeesOverride : [];
  if (employees.length === 0 || rows.length === 0) return rows;

  return rows.map((row, index) => {
    const emp = employees[index];
    if (!emp) return row;
    const fromEmp = mapFormAGJGujaratRowsFromEmployees([emp], hdrs, {
      ...helpers,
      rowIndexOffset: index,
    })[0];
    const out = { ...row };
    hdrs.forEach((header) => {
      const existing = out[header];
      const empVal = fromEmp?.[header];
      if (
        (existing == null || String(existing).trim() === '') &&
        empVal != null &&
        String(empVal).trim() !== ''
      ) {
        out[header] = empVal;
      }
    });
    if (fromEmp?.__employeeLookupName) out.__employeeLookupName = fromEmp.__employeeLookupName;
    return out;
  });
}

/** Write employee rows into the Gujarat Form A template without breaking column layout. */
export async function buildFormAGJGujaratWorkbookWithTemplateStyles({
  templateArrayBuffer,
  mappedData,
  headersToUse,
  parsedFormHeader,
  formFileName,
  headerFormData,
  parsedHeaderRowIndex,
  parsedDataStartIndex,
  parsedTableStartCol,
  employeesOverride = null,
  formatStatutoryDateDisplay = null,
}) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(templateArrayBuffer);
  const worksheet = workbook.worksheets[0];
  if (!worksheet) throw new Error('Template worksheet not found.');

  const layout = detectFormAGJGujaratTableLayout(worksheet, {
    parsedHeaderRowIndex,
    parsedDataStartIndex,
    parsedTableStartCol,
  });
  if (!layout) throw new Error('Could not locate Gujarat Form A table header row.');

  const { dataStartRow, templateCols } = layout;

  if (headerFormData && typeof headerFormData === 'object') {
    writeStatutoryHeaderFieldsToExcelJsWorksheet(worksheet, {
      headerFormData,
      parsedFormHeader,
      headerRowEnd: Math.max(1, layout.headerRow - 1),
      maxScanCols: 80,
    });
  }

  const templateHeaderLabels = templateCols.map((entry) => entry.label);
  const normalizedHeaders = resolveFormAGJGujaratTableHeaders(
    Array.isArray(headersToUse) && headersToUse.length >= templateHeaderLabels.length
      ? headersToUse
      : templateHeaderLabels.length > 0
        ? templateHeaderLabels
        : headersToUse
  );
  const exportHelpers = {
    formatStatutoryDateDisplay:
      typeof formatStatutoryDateDisplay === 'function'
        ? formatStatutoryDateDisplay
        : (v) => String(v || '').trim(),
  };
  const rows = enrichFormAGJGujaratRowsForExport(
    mappedData,
    normalizedHeaders,
    null,
    exportHelpers
  );

  const tableColMin = Math.min(...templateCols.map(({ col }) => col));
  const tableColMax = Math.max(...templateCols.map(({ col }) => col));
  const clearToRow = Math.max(dataStartRow + rows.length + 10, dataStartRow + 15);
  for (let r = dataStartRow; r <= clearToRow; r += 1) {
    for (let c = tableColMin; c <= tableColMax; c += 1) {
      worksheet.getCell(r, c).value = '';
    }
  }

  rows.forEach((row, idx) => {
    const targetRow = worksheet.getRow(dataStartRow + idx);
    if (targetRow) targetRow.height = 18;
    templateCols.forEach(({ col, bucket, label }) => {
      let val = getFormAGJGujaratRowValueForHeader(row, label, idx);
      if ((val == null || val === '') && bucket) {
        for (const [k, v] of Object.entries(row)) {
          if (
            formAGJHeaderAliasBucket(formAGJGujaratHeaderNorm(k)) === bucket &&
            v != null &&
            String(v).trim() !== ''
          ) {
            val = String(v).trim();
            break;
          }
        }
      }
      if ((val == null || val === '') && bucket === 'sno') val = String(idx + 1);
      const cell = worksheet.getCell(dataStartRow + idx, col);
      if (val == null || val === '') {
        cell.value = '';
        return;
      }
      if (bucket === 'sno') {
        const n = Number(String(val).replace(/[,]/g, '').trim());
        cell.value = Number.isFinite(n) ? n : String(val);
      } else if (bucket === 'mobile' || bucket === 'bankAccount' || bucket === 'aadhaar') {
        const digits = String(val).replace(/\D/g, '');
        cell.value = digits || String(val);
        if (digits) cell.numFmt = '0';
      } else {
        cell.value = String(val);
      }
      cell.alignment = { ...(cell.alignment || {}), vertical: 'middle', wrapText: true };
    });
  });

  if (rows.length > 0) {
    ensureExcelJSDataRowsWithBorders(worksheet, {
      dataStartRow,
      dataRowCount: rows.length,
      colFrom: tableColMin,
      colTo: tableColMax,
      templateRow: dataStartRow,
      templateBodyRows: 1,
    });
  }

  // Download: drop SCHEDULE / FORM A / FORMAT heading band; keep establishment + table.
  clearFormAGJGujaratExportTitleBands(worksheet);

  const out = await workbook.xlsx.writeBuffer();
  const fileName =
    formFileName ||
    parsedFormHeader?.title?.replace(/[^a-zA-Z0-9]/g, '_') ||
    'Form_A_GJ_Gujarat.xlsx';
  return {
    blob: new Blob([out], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    }),
    buffer: out,
    fileName,
  };
}
