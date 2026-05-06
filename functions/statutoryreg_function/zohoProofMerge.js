'use strict';

/**
 * Zoho People → statutory Excel proof template merge helpers.
 * Field mapping aligns with app/src/Pages/Statutory.js (fetchAndPopulateEmployeeData).
 */

const axios = require('axios');

const normalizeCellValue = (value) => String(value == null ? '' : value).trim();

async function getZohoAccessToken() {
  const refreshToken = process.env.ZOHO_REFRESH_TOKEN || '1000.22fee45e7bc3c714c7bd37fb5ffc6f3d.356011967bd84fb2e3051f79a6ed9779';
  const clientId = process.env.ZOHO_CLIENT_ID || '1000.L3KUHOMGRCUXORANE08T7B1HAUL0MP';
  const clientSecret = process.env.ZOHO_CLIENT_SECRET || '4579b2a1a3e19fa850eaae4c69691d78bbffe952e';

  if (!refreshToken || !clientId || !clientSecret) {
    throw new Error('Missing Zoho OAuth environment variables (ZOHO_REFRESH_TOKEN, ZOHO_CLIENT_ID, ZOHO_CLIENT_SECRET).');
  }

  const params = new URLSearchParams({
    refresh_token: refreshToken,
    client_id: clientId,
    client_secret: clientSecret,
    grant_type: 'refresh_token'
  });

  const { data } = await axios.post('https://accounts.zoho.in/oauth/v2/token', params);
  if (!data.access_token) {
    throw new Error('Failed to obtain access token from Zoho.');
  }
  return data.access_token;
}

async function fetchZohoEmployeeApi(limit = 200) {
  const accessToken = await getZohoAccessToken();
  const base = process.env.ZOHO_PEOPLE_BASE_URL || 'https://people.zoho.in/people/api';
  const endpoint = `${base}/forms/employee/getRecords`;
  const { data } = await axios.get(endpoint, {
    headers: { Authorization: `Bearer ${accessToken}` },
    params: { limit: String(limit) }
  });
  return data;
}

function flattenEmployeesFromZoho(data) {
  let employees = [];
  if (!data) return employees;

  if (data.response && data.response.result) {
    const result = data.response.result;
    if (Array.isArray(result)) {
      result.forEach((resultItem) => {
        if (resultItem && typeof resultItem === 'object') {
          Object.keys(resultItem).forEach((idKey) => {
            const employeeArray = resultItem[idKey];
            if (Array.isArray(employeeArray)) {
              employees.push(...employeeArray);
            } else if (employeeArray && typeof employeeArray === 'object') {
              employees.push(employeeArray);
            }
          });
        }
      });
    } else if (typeof result === 'object') {
      const firstKey = Object.keys(result)[0];
      if (firstKey && Array.isArray(result[firstKey])) {
        employees = result[firstKey];
      } else {
        employees = [result];
      }
    }
  } else if (Array.isArray(data)) {
    employees = data;
  } else if (data.records) {
    employees = data.records;
  }

  return employees;
}

function fallbackHeaderRowFromRows(rows) {
  const rowMeta = rows.map((row) => {
    const values = (row || []).map((cell) => normalizeCellValue(cell));
    const nonEmptyCount = values.filter(Boolean).length;
    return { values, nonEmptyCount };
  });
  let tableHeaderIndex = -1;
  let bestScore = -1;
  for (let i = 0; i < rowMeta.length; i += 1) {
    const current = rowMeta[i];
    if (!current || current.nonEmptyCount === 0) continue;
    const nextHasData = rowMeta[i + 1]?.nonEmptyCount > 0;
    if (!nextHasData) continue;
    const prevEmpty = i === 0 || rowMeta[i - 1]?.nonEmptyCount === 0;
    const score = current.nonEmptyCount + (prevEmpty ? 2 : 0);
    if (score > bestScore) {
      bestScore = score;
      tableHeaderIndex = i;
    }
  }
  if (tableHeaderIndex >= 0) {
    return {
      headerRowIndex: tableHeaderIndex,
      tableHeaders: rowMeta[tableHeaderIndex].values.map((c) => normalizeCellValue(c))
    };
  }
  return { headerRowIndex: -1, tableHeaders: [] };
}

/**
 * Worksheet that contains the employee register row (often "FORM U", not SheetNames[0]).
 */
function locateProofTemplateTable(workbook, XLSX) {
  for (const sheetName of workbook.SheetNames) {
    const sheet = workbook.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
    for (let i = 0; i < rows.length; i += 1) {
      const row = rows[i] || [];
      const lowered = row.map((c) => normalizeCellValue(c).toLowerCase());
      const joined = lowered.join(' ');
      if (
        (joined.includes('s.no') || joined.includes('s no') || joined.includes('serial')) &&
        joined.includes('name') &&
        (joined.includes('employee') || joined.includes('worker'))
      ) {
        return {
          sheetName,
          headerRowIndex: i,
          tableHeaders: row.map((c) => normalizeCellValue(c)),
          rows
        };
      }
    }
  }
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
  const fb = fallbackHeaderRowFromRows(rows);
  return {
    sheetName,
    headerRowIndex: fb.headerRowIndex,
    tableHeaders: fb.tableHeaders,
    rows
  };
}

/**
 * Detect employee table header row in template (e.g. FORM U row with S.No + Name of the employee).
 */
function extractTemplateTableHeaders(templateBuffer, XLSX) {
  const workbook = XLSX.read(templateBuffer, { type: 'buffer' });
  const loc = locateProofTemplateTable(workbook, XLSX);
  return loc.tableHeaders;
}

function mapOneField(emp, header, index) {
  const headerLower = header.toLowerCase().trim();

  if (headerLower.includes('s.no') || headerLower.includes('serial') || headerLower === 's.no') {
    return String(index + 1);
  }

  if (headerLower.includes('name') && (headerLower.includes('employee') || headerLower.includes('emp') || headerLower.includes('worker'))) {
    return (
      emp.FirstName ||
      emp.First_Name ||
      emp['First Name'] ||
      emp.Name1 ||
      emp.Nameoftheemployee ||
      emp['Name of the employee'] ||
      emp.name ||
      emp.employeeName ||
      emp.Name ||
      emp['Employee Name'] ||
      emp.Full_Name ||
      emp['Full Name'] ||
      emp.Employee_Name ||
      ''
    );
  }

  // "Employee Identific" and similar truncated Form U headers
  if (
    headerLower.includes('identific') ||
    headerLower.includes('identification') ||
    (headerLower.includes('employee') && headerLower.includes('id') && !headerLower.includes('aadhaar')) ||
    (headerLower.includes('emp') && headerLower.includes('code'))
  ) {
    return (
      emp.EmployeeID ||
      emp['EmployeeID'] ||
      emp['Role.ID'] ||
      (emp.Role && typeof emp.Role === 'object' ? emp.Role.ID : null) ||
      emp.EmployeeIdentificationNo ||
      emp['Employee Identification No.'] ||
      emp.employeeCode ||
      emp.EmployeeCode ||
      emp['Employee Code'] ||
      emp['Employee ID'] ||
      emp.ID ||
      ''
    );
  }

  if ((headerLower.includes('worker') && headerLower.includes('identity')) || headerLower.includes('worker identity')) {
    return emp.Zoho_ID || emp.ZohoID || emp.zoho_id || emp.zohoId || '';
  }

  if (headerLower === 'emp id' || headerLower === 'empid' || (headerLower.includes('emp') && headerLower.includes('id') && !headerLower.includes('identification'))) {
    return emp.Zoho_ID || emp.ZohoID || emp.zoho_id || emp.zohoId || '';
  }

  if (headerLower.includes('gender') || headerLower.includes('sex')) {
    return emp.Sex || emp.Gender || emp.gender || '';
  }

  if (headerLower.includes('father') || headerLower.includes('spouse')) {
    const fatherName =
      emp.Father_s_Name ||
      emp.Father_Name ||
      emp.FathersName ||
      emp['Father/Spouse Name'] ||
      emp['Father Name'] ||
      '';
    const spouseName = emp.Spouse_Name || emp.spouseName || emp['Spouse Name'] || '';
    return (fatherName && String(fatherName).trim()) || (spouseName && String(spouseName).trim()) ? String(fatherName || spouseName).trim() : '';
  }

  if (headerLower.includes('date') && headerLower.includes('birth')) {
    return (
      emp.Date_of_birth ||
      emp.DateofBirth ||
      emp['Date of Birth'] ||
      emp.dateOfBirth ||
      emp.Date_Of_Birth ||
      ''
    );
  }

  if (headerLower.includes('date') && headerLower.includes('join')) {
    return (
      emp.Dateofjoining ||
      emp.DateofJoining ||
      emp['Date of Joining'] ||
      emp.dateOfJoining ||
      emp.Date_of_Joining ||
      ''
    );
  }

  if (headerLower.includes('date') && headerLower.includes('permanent') && headerLower.includes('made')) {
    return emp.Dateofjoining || emp['Dateofjoining'] || emp['Date of Joining'] || '';
  }

  if (headerLower.includes('designation')) {
    return emp.Designation || emp.designation || '';
  }

  if (headerLower.includes('present') && headerLower.includes('address')) {
    const addressLine1 = emp.Address_Line_1 || emp['Address_Line_1'] || '';
    const country = emp.Country || '';
    const city1 = emp.City1 || '';
    const city = emp.City || '';
    const addressParts = [];
    if (addressLine1 && String(addressLine1).trim()) addressParts.push(String(addressLine1).trim());
    if (city1 && String(city1).trim()) addressParts.push(String(city1).trim());
    if (city && String(city).trim() && String(city).trim() !== String(city1).trim()) addressParts.push(String(city).trim());
    if (!city1 && city && String(city).trim()) addressParts.push(String(city).trim());
    if (country && String(country).trim()) addressParts.push(String(country).trim());
    return addressParts.length ? addressParts.join(', ') : emp.PresentAddress || emp['Present Address'] || '';
  }

  if (headerLower.includes('permanent') && headerLower.includes('address')) {
    const addressLine1 = emp.Address_Line_1 || emp['Address_Line_1'] || '';
    const addressLine2 = emp.Address_Line_2 || emp['Address_Line_2'] || '';
    const country1 = emp.Country1 || '';
    const city = emp.City || '';
    const city2 = emp.City2 || '';
    const addressParts = [];
    if (addressLine1 && String(addressLine1).trim()) addressParts.push(String(addressLine1).trim());
    if (addressLine2 && String(addressLine2).trim()) addressParts.push(String(addressLine2).trim());
    if (city && String(city).trim()) addressParts.push(String(city).trim());
    if (city2 && String(city2).trim() && String(city2).trim() !== String(city).trim()) addressParts.push(String(city2).trim());
    if (!city && city2 && String(city2).trim()) addressParts.push(String(city2).trim());
    if (country1 && String(country1).trim()) addressParts.push(String(country1).trim());
    return addressParts.length ? addressParts.join(', ') : emp.Permanentaddress || emp['Permanent address'] || '';
  }

  if (headerLower.includes('provident') || (headerLower.includes('pf') && headerLower.includes('no'))) {
    const pfValue =
      emp["Employee's Provident Fund No."] ||
      emp.ProvidentFundNo ||
      emp['Provident_Fund_No'] ||
      emp.PF_No ||
      '';
    return pfValue && String(pfValue).trim() && !String(pfValue).toLowerCase().includes('zoho') ? String(pfValue).trim() : '';
  }

  if (headerLower.includes('esic') || (headerLower.includes('insurance') && headerLower.includes('state'))) {
    return emp['Employee\'s State Insurance Corporation No.'] || emp.ESICNo || emp.esicNo || '';
  }

  if (headerLower.includes('aadhaar') || headerLower.includes('aadhar')) {
    return emp.AadhaarNo || emp['Aadhaar No.'] || '';
  }

  if (headerLower.includes('email')) {
    return emp.EmailID || emp.Email || emp.email || emp['Email ID'] || '';
  }

  if (headerLower.includes('mobile') || (headerLower.includes('phone') && !headerLower.includes('emergency'))) {
    return (
      emp.Mobile ||
      emp.Mobile_Number ||
      emp['Mobile Number'] ||
      emp.Phone ||
      emp.Phone_Number ||
      ''
    );
  }

  if (headerLower.includes('department')) {
    return emp.Department || emp.department || '';
  }

  if (
    headerLower.includes('bank') &&
    headerLower.includes('ifsc') &&
    (headerLower.includes('account') || headerLower.includes('name') || headerLower.includes('branch'))
  ) {
    const accountNumber = emp.Account_Number || emp.AccountNumber || '';
    const bankName = emp.Bank_Name || emp.BankName || '';
    const branchName = emp.Branch_Name || emp.BranchName || '';
    const ifscCode = emp.IFSC_Code || emp.IFSCCode || '';
    const bankParts = [];
    if (accountNumber && String(accountNumber).trim()) bankParts.push(String(accountNumber).trim());
    if (bankName && String(bankName).trim()) bankParts.push(String(bankName).trim());
    if (branchName && String(branchName).trim()) bankParts.push(String(branchName).trim());
    let combined = bankParts.length ? bankParts.join(', ') : '';
    if (ifscCode && String(ifscCode).trim()) {
      combined = combined ? `${combined} (${String(ifscCode).trim()})` : String(ifscCode).trim();
    }
    return combined;
  }

  if (headerLower.includes('ifsc') && !headerLower.includes('bank')) {
    return emp.IFSC_Code || emp.IFSCCode || emp.IFSC || '';
  }

  if (headerLower.includes('branch') && !headerLower.includes('ifsc') && !(headerLower.includes('bank') && headerLower.includes('account'))) {
    return emp.Branch_Name || emp.BranchName || '';
  }

  if (headerLower.includes('bank') && (headerLower.includes('account') || headerLower.includes('ac')) && !headerLower.includes('name')) {
    return emp.Account_Number || emp.AccountNumber || '';
  }

  if (headerLower.includes('name') && headerLower.includes('bank') && !headerLower.includes('account')) {
    return emp.Bank_Name || emp.BankName || '';
  }

  if (headerLower.includes('reason') && headerLower.includes('exit')) {
    return emp.Reason_for_Exit || emp['Reason for Exit'] || '';
  }

  if (headerLower.includes('date') && headerLower.includes('left')) {
    return emp.Dateofexit || emp['Date of Exit'] || emp.LastWorkingDate || '';
  }

  return '';
}

function mapZohoEmployeeToRow(empItem, headers, index) {
  const emp = empItem.Employee || empItem.employee || empItem;
  const row = {};
  headers.forEach((header) => {
    if (!normalizeCellValue(header)) return;
    row[header] = mapOneField(emp, header, index);
  });
  return row;
}

/**
 * @returns {Promise<{ tableHeaders: string[], tableData: object[], headerFields: object }|null>}
 */
async function buildProofDataFromZoho(templateBuffer, XLSX, options = {}) {
  const limit = options.limit || 200;
  const apiData = await fetchZohoEmployeeApi(limit);
  const employees = flattenEmployeesFromZoho(apiData);
  const workbook = XLSX.read(templateBuffer, { type: 'buffer' });
  const loc = locateProofTemplateTable(workbook, XLSX);

  if (
    loc.headerRowIndex < 0 ||
    !loc.tableHeaders.some((h) => normalizeCellValue(h)) ||
    !employees.length
  ) {
    return null;
  }

  const tableData = employees.map((e, i) => mapZohoEmployeeToRow(e, loc.tableHeaders, i));
  return {
    tableHeaders: loc.tableHeaders,
    tableData,
    headerFields: {},
    sheetName: loc.sheetName,
    tableHeaderRowIndex: loc.headerRowIndex
  };
}

module.exports = {
  buildProofDataFromZoho,
  flattenEmployeesFromZoho,
  extractTemplateTableHeaders,
  fetchZohoEmployeeApi,
  normalizeCellValue,
  locateProofTemplateTable
};
