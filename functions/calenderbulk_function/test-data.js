// Test data for Calendar Bulk function
// This matches the structure expected by the Data Store schema

const testCalendarBulkData = [
  {
    sNo: "1",
    forms: "Form 3",
    sector: "Manufacturing",
    description: "Construction Permission",
    rules: "Rule 3",
    schedule: "Permission to be obtained with relevant documents before construction, reconstruction or extension.",
    jan: "",
    feb: "",
    mar: "",
    apr: "",
    may: "",
    june: "",
    july: "",
    aug: "31",
    sep: "",
    oct: "",
    nov: "",
    dec: ""
  },
  {
    sNo: "2",
    forms: "Form 4",
    sector: "Manufacturing",
    description: "Registration/License Renewal",
    rules: "Rule 4",
    schedule: "Registration/license to be obtained before commencement of manufacturing process for the first time and application to be submitted on or before October 31st of every year for renewal.",
    jan: "",
    feb: "",
    mar: "",
    apr: "",
    may: "",
    june: "",
    july: "",
    aug: "",
    sep: "",
    oct: "31",
    nov: "",
    dec: ""
  },
  {
    sNo: "3",
    forms: "Form 3-A",
    sector: "Healthcare",
    description: "Manager Change Notice",
    rules: "Rule 12-A",
    schedule: "The notice of change of manager shall be in Form 3-A.",
    jan: "",
    feb: "",
    mar: "",
    apr: "",
    may: "",
    june: "",
    july: "",
    aug: "",
    sep: "",
    oct: "31",
    nov: "",
    dec: ""
  },
  {
    sNo: "4",
    forms: "Form 4(6)",
    sector: "Finance",
    description: "Registration Authority",
    rules: "Rule 4(6)",
    schedule: "The registration and license shall be obtained from the concerned authority.",
    jan: "",
    feb: "",
    mar: "",
    apr: "",
    may: "",
    june: "",
    july: "",
    aug: "",
    sep: "",
    oct: "",
    nov: "",
    dec: "15"
  },
  {
    sNo: "78",
    forms: "Form 10",
    sector: "Manufacturing",
    description: "Overtime Work Entry",
    rules: "Rule 78",
    schedule: "Period of overtime work shall be entered in Form 10.",
    jan: "",
    feb: "",
    mar: "",
    apr: "",
    may: "",
    june: "",
    july: "",
    aug: "",
    sep: "30",
    oct: "",
    nov: "",
    dec: ""
  },
  {
    sNo: "80 & 86",
    forms: "Form 12",
    sector: "Manufacturing",
    description: "Register of Adult Workers",
    rules: "Rule 80 & 86",
    schedule: "The Register of adult workers shall be in Form 12.",
    jan: "",
    feb: "",
    mar: "",
    apr: "",
    may: "",
    june: "",
    july: "",
    aug: "",
    sep: "30",
    oct: "",
    nov: "",
    dec: ""
  },
  {
    sNo: "87 & 88",
    forms: "Form 15",
    sector: "Manufacturing",
    description: "Leave Register",
    rules: "Rule 87 & 88",
    schedule: "The manager shall keep an up-to-date Register in Form 15 for leave",
    jan: "",
    feb: "",
    mar: "",
    apr: "",
    may: "",
    june: "",
    july: "",
    aug: "",
    sep: "30",
    oct: "",
    nov: "",
    dec: ""
  },
  {
    sNo: "96",
    forms: "Accident Report",
    sector: "Manufacturing",
    description: "Accident Reporting",
    rules: "Rule 96",
    schedule: "Within 12 hours for death or dangerous occurrence and within 24 hours of the expiry of 48 hours of the accident (bodily injury), the factory manager shall send the report to the concerned authority.",
    jan: "",
    feb: "",
    mar: "",
    apr: "",
    may: "",
    june: "",
    july: "",
    aug: "",
    sep: "30",
    oct: "",
    nov: "",
    dec: ""
  }
];

// Sample request objects for testing different API endpoints
const testRequests = {
  // Test getAll action
  getAll: {
    method: 'GET',
    query: { action: 'getAll' },
    body: null
  },
  
  // Test getById action
  getById: {
    method: 'GET',
    query: { action: 'getById', id: '1' },
    body: null
  },
  
  // Test add action
  add: {
    method: 'POST',
    query: { action: 'add' },
    body: testCalendarBulkData[0]
  },
  
  // Test bulkImport action
  bulkImport: {
    method: 'POST',
    query: { action: 'bulkImport' },
    body: testCalendarBulkData
  },
  
  // Test getBySector action
  getBySector: {
    method: 'GET',
    query: { action: 'getBySector', sector: 'Manufacturing' },
    body: null
  },
  
  // Test count action
  count: {
    method: 'GET',
    query: { action: 'count' },
    body: null
  },
  
  // Test update action
  update: {
    method: 'PUT',
    query: { action: 'update', id: '1' },
    body: {
      ...testCalendarBulkData[0],
      description: 'Updated Construction Permission Description'
    }
  },
  
  // Test delete action
  delete: {
    method: 'DELETE',
    query: { action: 'delete', id: '1' },
    body: null
  }
};

module.exports = {
  testCalendarBulkData,
  testRequests
};
