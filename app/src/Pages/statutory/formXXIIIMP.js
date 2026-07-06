import { readPayrollNetPayForStatutory } from '../../utils/payrollEarnings';
import { isAprilPayrollMonthCandidates } from './formQKarnataka';
import { personNamesMatch } from './formFKarnataka';

/** Form XXIII MP — Register of Overtime (Madhya Pradesh). */

/** April default Normal rate of wages from Form XXIII_MP template when payroll lacks net pay. */
export const FORM_XXIII_MP_APR_DEFAULT_PAYROLL = [
  { name: 'Prem Singh Bhati', normalRate: '122603' },
  { name: 'Satyapriya Behera', normalRate: '92766' },
  { name: 'Dhavalkumar Hiteshbhai Gondaliya', normalRate: '137383' },
  { name: 'Rahul Kushwah', normalRate: '112053' },
  { name: 'Katariya Jitendra', normalRate: '118855' },
  { name: 'Shekhar Sharma', normalRate: '111112' },
  { name: 'Rajan Kumar Keshari', normalRate: '179788' },
  { name: 'Manoj Kumar Gurjar', normalRate: '57916' },
  { name: 'Nil Kamal Sahu', normalRate: '63398' },
  { name: 'Amit Kumar', normalRate: '143114' },
  { name: 'Niranjan Sahoo', normalRate: '98051' },
  { name: 'Rajanish Kumar Maurya', normalRate: '100079' },
  { name: 'Amartya Pal', normalRate: '112105' },
  { name: 'Sudhir Kumar Raula', normalRate: '64768' },
  { name: 'Prakash Hamirbhai Chudasama', normalRate: '113958' },
  { name: 'Anuj Rana', normalRate: '109046' },
  { name: 'Rakesh Behera', normalRate: '100893' },
  { name: 'Harshad Dk', normalRate: '109329' },
  { name: 'Anuj Kumar', normalRate: '100079' },
  { name: 'Vivek Pandey', normalRate: '113958' },
  { name: 'Hardik Nandaniya', normalRate: '111182' },
  { name: 'Manjunatha Yarekoppa', normalRate: '144760' },
  { name: 'Darshit Patel', normalRate: '92977' },
  { name: 'Pradeep Kumar Nayak', normalRate: '101311' },
  { name: 'Jagdish Vasan', normalRate: '14036' },
  { name: 'Pankaj Kushwah', normalRate: '98140' },
  { name: 'Mohalkhram B', normalRate: '94925' },
  { name: 'Narendra Yogi', normalRate: '96515' },
  { name: 'Susanta Rout', normalRate: '64901' },
  { name: 'Rohit Kumar', normalRate: '84650' },
  { name: 'Manikandan V', normalRate: '87217' },
  { name: 'Shrinath Shukla', normalRate: '110861' },
  { name: 'Aditya Sharma', normalRate: '90134' },
  { name: 'Harish Chandra Pandey Harish Chandra Pandey', normalRate: '101533' },
  { name: 'Rahul Pal', normalRate: '92317' },
  { name: 'Vikram Sharma', normalRate: '109405' },
  { name: 'Mannu Singh', normalRate: '103837' },
  { name: 'Sartanbhai Parmabhai Pagi', normalRate: '90454' },
  { name: 'Ashwin Vikramshi Suva', normalRate: '85085' },
  { name: 'Rahul Patil', normalRate: '99829' },
  { name: 'Jayprakash Patra', normalRate: '94343' },
  { name: 'Janak Mukeshbhai Bhaliya', normalRate: '97999' },
  { name: 'Himansu Sekhar Panda', normalRate: '93185' },
  { name: 'Karuppasamy Erulappan', normalRate: '68686' },
  { name: 'Sampat Ram Unkaram', normalRate: '64724' },
  { name: 'Pradeep Saini Durga Prasad', normalRate: '99157' },
  { name: 'Piyush Kanjariya Odhavajibhai', normalRate: '120728' },
  { name: 'Maheshbhai Meda Ratnabhai Kuberbhai', normalRate: '96849' },
  { name: 'Jaydeep Dineshbhai Dineshbhai', normalRate: '115928' },
  { name: 'Ajaykumar Mansingbhai', normalRate: '102165' },
  { name: 'Dharmendra Kumar Yadav Haripal', normalRate: '138836' },
  { name: 'Patel Bhavin Kodarbhai', normalRate: '89710' },
  { name: 'Patel Chandrakant Virabhai', normalRate: '98551' },
  { name: 'Chandrakant Parmar', normalRate: '114766' },
  { name: 'Sujit Kumar Shivanna', normalRate: '121170' },
  { name: 'Ravi Kumar Singh', normalRate: '106452' },
  { name: 'Rajeshkumar Jamabhai Nai', normalRate: '117217' },
  { name: 'Rajan Vijay Arote', normalRate: '147131' },
  { name: 'Sachin Sharma', normalRate: '158629' },
  { name: 'Pradeep Sampath', normalRate: '87314' },
  { name: 'Muthukumar Paramasivan', normalRate: '209018' },
  { name: 'Kolappan Madevan Pillai', normalRate: '266963' },
];

export function formatFormXXIIIMPEmployeeName(emp = {}) {
  const fn = String(emp.FirstName || emp['FirstName'] || emp.firstName || emp['First Name'] || '').trim();
  const ln = String(emp.LastName || emp['LastName'] || emp.lastName || emp['Last Name'] || '').trim();
  if (fn && ln) return `${fn} ${ln}`;
  return fn || ln || String(emp.Name || emp['Name'] || emp.EmployeeName || emp['Employee Name'] || '').trim();
}

export function resolveFormXXIIIMPAprilDefaultNormalRate(emp) {
  const empName = formatFormXXIIIMPEmployeeName(emp);
  if (!empName) return '';
  const match = FORM_XXIII_MP_APR_DEFAULT_PAYROLL.find((entry) => personNamesMatch(empName, entry.name));
  return match?.normalRate != null ? String(match.normalRate).trim() : '';
}

export function isFormXXIIIMPContext(formHeader, rowItem, fileName, sheetText = '') {
  const parts = [
    rowItem?.formName,
    rowItem?.FormName,
    rowItem?.state,
    rowItem?.State,
    rowItem?.siteState,
    rowItem?.SiteState,
    rowItem?.formFileName,
    rowItem?.FormFileName,
    fileName,
    formHeader?.title,
    formHeader?.subtitle,
    sheetText,
  ]
    .filter((x) => x != null && String(x).trim() !== '')
    .join(' ')
    .toLowerCase();

  const hasMP =
    /madhya[\s._-]*pradesh/.test(parts) ||
    /\bform[\s._-]*xxiii[\s._-]*mp\b/.test(parts) ||
    /\bxxiii[\s._-]*mp\b/.test(parts);
  const hasXXIII =
    /\bform[\s._-]*xxiii/i.test(parts) || /register\s+of\s+overtime/.test(parts);
  return hasMP && hasXXIII;
}

/** Normal rate of wages — April template defaults, then payroll net pay. */
export function resolveFormXXIIIMPNormalRateForEmployee(emp, payrollRow = null, monthCandidates = null) {
  if (isAprilPayrollMonthCandidates(monthCandidates)) {
    const aprilDefault = resolveFormXXIIIMPAprilDefaultNormalRate(emp);
    if (aprilDefault) return aprilDefault;
  }
  if (payrollRow && !payrollRow.fetch_error) {
    return String(readPayrollNetPayForStatutory(payrollRow) ?? '').trim();
  }
  return '';
}
