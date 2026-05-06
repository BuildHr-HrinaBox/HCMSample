'use strict';

const express = require('express');
const catalystSDK = require('zcatalyst-sdk-node');

const app = express();

app.use(express.json());

// Attach catalyst instance per request
app.use((req, res, next) => {
  try {
    const catalyst = catalystSDK.initialize(req);
    res.locals.catalyst = catalyst;
    next();
  } catch (err) {
    res.status(500).send({ status: 'failure', message: 'Catalyst init failed' });
  }
});

// Table name for calendar bulk data
const TABLE_NAME = 'Calenderbulk';

// Calendar bulk data structure based on the Data Store schema
const createCalendarBulkRecord = (data) => {
  return {
    Forms: data.forms || '',
    Sector: data.sector || '',
    Description: data.description || '',
    Rules: data.rules || '',
    ScheduleOfSubmission: data.schedule || '',
    Jan: data.jan || '',
    Feb: data.feb || '',
    Mar: data.mar || '',
    Apr: data.apr || '',
    May: data.may || '',
    June: data.june || '',
    July: data.july || '',
    Aug: data.aug || '',
    Sep: data.sep || '',
    Oct: data.oct || '',
    Nov: data.nov || '',
    December: data.dec || ''
  };
};

// Convert Data Store record back to application format
const convertToAppFormat = (record) => {
  return {
    id: record.ROWID,
    sNo: record.SNo || '',
    forms: record.Forms || '',
    sector: record.Sector || '',
    description: record.Description || '',
    rules: record.Rules || '',
    schedule: record.ScheduleOfSubmission || '',
    jan: record.Jan || '',
    feb: record.Feb || '',
    mar: record.Mar || '',
    apr: record.Apr || '',
    may: record.May || '',
    june: record.June || '',
    july: record.July || '',
    aug: record.Aug || '',
    sep: record.Sep || '',
    oct: record.Oct || '',
    nov: record.Nov || '',
    dec: record.December || ''
  };
};

// Get all calendar bulk data
const getAllCalendarBulkData = async (catalyst) => {
  try {
    console.log('Fetching all calendar bulk data from Data Store...');
    
    const dataStore = catalyst.datastore();
    const table = dataStore.table(TABLE_NAME);
    
    // Use getAllRows method instead of select
    const result = await table.getAllRows();
    
    console.log(`Found ${result.length} calendar bulk records`);
    
    const formattedData = result.map(record => convertToAppFormat(record));
    
    return {
      status: 'success',
      message: 'Calendar bulk data retrieved successfully',
      data: formattedData,
      count: formattedData.length
    };
  } catch (error) {
    console.error('Error fetching calendar bulk data:', error);
    return {
      status: 'error',
      message: 'Failed to fetch calendar bulk data',
      error: error.message
    };
  }
};

// Get calendar bulk data by ID
const getCalendarBulkById = async (catalyst, id) => {
  try {
    console.log(`Fetching calendar bulk data for ID: ${id}`);
    
    const dataStore = catalyst.datastore();
    const table = dataStore.table(TABLE_NAME);
    
    // Use getRow method instead of select
    const result = await table.getRow(id);
    
    if (!result) {
      return {
        status: 'error',
        message: 'Calendar bulk record not found'
      };
    }
    
    const formattedData = convertToAppFormat(result);
    
    return {
      status: 'success',
      message: 'Calendar bulk data retrieved successfully',
      data: formattedData
    };
  } catch (error) {
    console.error('Error fetching calendar bulk data by ID:', error);
    return {
      status: 'error',
      message: 'Failed to fetch calendar bulk data',
      error: error.message
    };
  }
};

// Add new calendar bulk data
const addCalendarBulkData = async (catalyst, data) => {
  try {
    console.log('Adding new calendar bulk data:', data);
    
    const dataStore = catalyst.datastore();
    const table = dataStore.table(TABLE_NAME);
    const record = createCalendarBulkRecord(data);
    
    const result = await table.insertRow(record);
    
    console.log('Calendar bulk data added successfully with ROWID:', result.ROWID);
    
    return {
      status: 'success',
      message: 'Calendar bulk data added successfully',
      data: {
        id: result.ROWID,
        ...data
      }
    };
  } catch (error) {
    console.error('Error adding calendar bulk data:', error);
    return {
      status: 'error',
      message: 'Failed to add calendar bulk data',
      error: error.message
    };
  }
};

// Update calendar bulk data
const updateCalendarBulkData = async (catalyst, id, data) => {
  try {
    console.log(`Updating calendar bulk data for ID: ${id}`, data);
    
    const dataStore = catalyst.datastore();
    const table = dataStore.table(TABLE_NAME);
    const record = createCalendarBulkRecord(data);
    
    const result = await table.updateRow(id, record);
    
    console.log('Calendar bulk data updated successfully');
    
    return {
      status: 'success',
      message: 'Calendar bulk data updated successfully',
      data: {
        id: id,
        ...data
      }
    };
  } catch (error) {
    console.error('Error updating calendar bulk data:', error);
    return {
      status: 'error',
      message: 'Failed to update calendar bulk data',
      error: error.message
    };
  }
};

// Delete calendar bulk data
const deleteCalendarBulkData = async (catalyst, id) => {
  try {
    console.log(`Deleting calendar bulk data for ID: ${id}`);
    
    const dataStore = catalyst.datastore();
    const table = dataStore.table(TABLE_NAME);
    await table.deleteRow(id);
    
    console.log('Calendar bulk data deleted successfully');
    
    return {
      status: 'success',
      message: 'Calendar bulk data deleted successfully'
    };
  } catch (error) {
    console.error('Error deleting calendar bulk data:', error);
    return {
      status: 'error',
      message: 'Failed to delete calendar bulk data',
      error: error.message
    };
  }
};

// Bulk import calendar data
const bulkImportCalendarData = async (catalyst, dataArray) => {
  try {
    console.log(`Bulk importing ${dataArray.length} calendar bulk records`);
    
    const dataStore = catalyst.datastore();
    const table = dataStore.table(TABLE_NAME);
    const records = dataArray.map(data => createCalendarBulkRecord(data));
    
    const result = await table.insertRows(records);
    
    console.log(`Bulk import completed. ${result.length} records added`);
    
    return {
      status: 'success',
      message: `Bulk import completed successfully. ${result.length} records added`,
      data: result.map((record, index) => ({
        id: record.ROWID,
        ...dataArray[index]
      }))
    };
  } catch (error) {
    console.error('Error in bulk import:', error);
    return {
      status: 'error',
      message: 'Failed to bulk import calendar data',
      error: error.message
    };
  }
};

// Get all unique sectors from calendar bulk data
const getAllUniqueSectors = async (catalyst) => {
  try {
    console.log('Fetching all unique sectors...');
    const table = catalyst.datastore().table(TABLE_NAME);
    
    // Get all rows
    const allRows = await table.getAllRows();
    console.log(`Retrieved ${allRows.length} total records`);
    
    // Extract unique sectors (case-insensitive)
    const sectors = [...new Set(allRows.map(record => record.Sector?.toLowerCase()).filter(sector => sector && sector.trim()))];
    console.log(`Found ${sectors.length} unique sectors:`, sectors);
    
    return {
      status: 'success',
      message: 'Unique sectors retrieved successfully',
      data: {
        sectors: sectors.sort() // Sort alphabetically
      }
    };
  } catch (error) {
    console.error('Error fetching unique sectors:', error);
    return {
      status: 'error',
      message: 'Failed to fetch unique sectors',
      error: error.message
    };
  }
};

// Get calendar data by sector/industry
const getCalendarDataBySector = async (catalyst, sector) => {
  try {
    console.log(`Fetching calendar data for sector: ${sector}`);
    
    const dataStore = catalyst.datastore();
    const table = dataStore.table(TABLE_NAME);
    
    // Get all rows and filter by sector
    const allRows = await table.getAllRows();
    const filteredRows = allRows.filter(record => 
      record.Sector && record.Sector.toLowerCase() === sector.toLowerCase()
    );
    
    console.log(`Found ${filteredRows.length} records for sector: ${sector}`);
    
    const formattedData = filteredRows.map(record => convertToAppFormat(record));
    
    return {
      status: 'success',
      message: `Calendar data for sector '${sector}' retrieved successfully`,
      data: formattedData,
      count: formattedData.length
    };
  } catch (error) {
    console.error('Error fetching calendar data by sector:', error);
    return {
      status: 'error',
      message: 'Failed to fetch calendar data by sector',
      error: error.message
    };
  }
};

// Get count of calendar bulk records
const getCalendarBulkCount = async (catalyst) => {
  try {
    console.log('Getting calendar bulk data count...');
    
    const dataStore = catalyst.datastore();
    const table = dataStore.table(TABLE_NAME);
    
    // Use getAllRows and get length
    const result = await table.getAllRows();
    
    return {
      status: 'success',
      message: 'Calendar bulk data count retrieved successfully',
      count: result.length
    };
  } catch (error) {
    console.error('Error getting calendar bulk count:', error);
    return {
      status: 'error',
      message: 'Failed to get calendar bulk data count',
      error: error.message
    };
  }
};

// Populate table with sample data
const populateTableWithSampleData = async (catalyst) => {
  try {
    console.log('Populating Calenderbulk table with sample data...');
    
    const sampleData = [
      {
        Forms: "Form 3",
        Sector: "Manufacturing",
        Description: "Construction Permission",
        Rules: "Rule 3",
        ScheduleOfSubmission: "Permission to be obtained with relevant documents before construction, reconstruction or extension.",
        Jan: "",
        Feb: "",
        Mar: "",
        Apr: "",
        May: "",
        June: "",
        July: "",
        Aug: "31",
        Sep: "",
        Oct: "",
        Nov: "",
        December: ""
      },
      {
        Forms: "Form 4",
        Sector: "Manufacturing",
        Description: "Registration/License Renewal",
        Rules: "Rule 4",
        ScheduleOfSubmission: "Registration/license to be obtained before commencement of manufacturing process for the first time and application to be submitted on or before October 31st of every year for renewal.",
        Jan: "",
        Feb: "",
        Mar: "",
        Apr: "",
        May: "",
        June: "",
        July: "",
        Aug: "",
        Sep: "",
        Oct: "31",
        Nov: "",
        December: ""
      },
      {
        Forms: "Form 3-A",
        Sector: "Healthcare",
        Description: "Manager Change Notice",
        Rules: "Rule 12-A",
        ScheduleOfSubmission: "The notice of change of manager shall be in Form 3-A.",
        Jan: "",
        Feb: "",
        Mar: "",
        Apr: "",
        May: "",
        June: "",
        July: "",
        Aug: "",
        Sep: "",
        Oct: "31",
        Nov: "",
        December: ""
      },
      {
        Forms: "Form 4(6)",
        Sector: "Finance",
        Description: "Registration Authority",
        Rules: "Rule 4(6)",
        ScheduleOfSubmission: "The registration and license shall be obtained from the concerned authority.",
        Jan: "",
        Feb: "",
        Mar: "",
        Apr: "",
        May: "",
        June: "",
        July: "",
        Aug: "",
        Sep: "",
        Oct: "",
        Nov: "",
        December: "15"
      },
      {
        Forms: "Form 10",
        Sector: "Manufacturing",
        Description: "Overtime Work Entry",
        Rules: "Rule 78",
        ScheduleOfSubmission: "Period of overtime work shall be entered in Form 10.",
        Jan: "",
        Feb: "",
        Mar: "",
        Apr: "",
        May: "",
        June: "",
        July: "",
        Aug: "",
        Sep: "30",
        Oct: "",
        Nov: "",
        December: ""
      },
      {
        Forms: "Form 12",
        Sector: "Manufacturing",
        Description: "Register of Adult Workers",
        Rules: "Rule 80 & 86",
        ScheduleOfSubmission: "The Register of adult workers shall be in Form 12.",
        Jan: "",
        Feb: "",
        Mar: "",
        Apr: "",
        May: "",
        June: "",
        July: "",
        Aug: "",
        Sep: "30",
        Oct: "",
        Nov: "",
        December: ""
      },
      {
        Forms: "Form 15",
        Sector: "Manufacturing",
        Description: "Leave Register",
        Rules: "Rule 87 & 88",
        ScheduleOfSubmission: "The manager shall keep an up-to-date Register in Form 15 for leave",
        Jan: "",
        Feb: "",
        Mar: "",
        Apr: "",
        May: "",
        June: "",
        July: "",
        Aug: "",
        Sep: "30",
        Oct: "",
        Nov: "",
        December: ""
      },
      {
        Forms: "Accident Report",
        Sector: "Manufacturing",
        Description: "Accident Reporting",
        Rules: "Rule 96",
        ScheduleOfSubmission: "Within 12 hours for death or dangerous occurrence and within 24 hours of the expiry of 48 hours of the accident (bodily injury), the factory manager shall send the report to the concerned authority.",
        Jan: "",
        Feb: "",
        Mar: "",
        Apr: "",
        May: "",
        June: "",
        July: "",
        Aug: "",
        Sep: "30",
        Oct: "",
        Nov: "",
        December: ""
      }
    ];
    
    const dataStore = catalyst.datastore();
    const table = dataStore.table(TABLE_NAME);
    const result = await table.insertRows(sampleData);
    
    console.log(`Successfully populated table with ${result.length} sample records`);
    
    return {
      status: 'success',
      message: `Successfully populated Calenderbulk table with ${result.length} sample records`,
      data: result.map(record => ({
        id: record.ROWID,
        forms: record.Forms,
        sector: record.Sector,
        description: record.Description
      }))
    };
  } catch (error) {
    console.error('Error populating table with sample data:', error);
    return {
      status: 'error',
      message: 'Failed to populate table with sample data',
      error: error.message
    };
  }
};

// Health check
app.get('/', (req, res) => {
  res.status(200).json({ status: 'success', message: 'calenderbulk_function ready' });
});

// Get all calendar bulk data
app.get('/calenderbulk', async (req, res) => {
  try {
    const catalyst = res.locals.catalyst;
    const { action, id, sector } = req.query;
    
    console.log(`Calendar Bulk API Request - Action: ${action}, ID: ${id}, Sector: ${sector}`);
    
    let result;
    
    switch (action) {
      case 'getAll':
        result = await getAllCalendarBulkData(catalyst);
        break;
        
      case 'getSectors':
        result = await getAllUniqueSectors(catalyst);
        break;
        
      case 'getById':
        if (!id) {
          result = {
            status: 'error',
            message: 'ID parameter is required for getById action'
          };
        } else {
          result = await getCalendarBulkById(catalyst, id);
        }
        break;
        
      case 'getBySector':
        if (!sector) {
          result = {
            status: 'error',
            message: 'Sector parameter is required for getBySector action'
          };
        } else {
          result = await getCalendarDataBySector(catalyst, sector);
        }
        break;
        
      case 'count':
        result = await getCalendarBulkCount(catalyst);
        break;
        
      case 'populate':
        result = await populateTableWithSampleData(catalyst);
        break;
        
      default:
        result = {
          status: 'error',
          message: 'Invalid action. Supported actions: getAll, getById, getBySector, getSectors, count, populate'
        };
    }
    
    res.status(result.status === 'success' ? 200 : 400).json(result);
    
  } catch (error) {
    console.error('Calendar Bulk API Error:', error);
    res.status(500).json({
      status: 'error',
      message: 'Internal server error',
      error: error.message
    });
  }
});

// POST routes for data modification
app.post('/calenderbulk', async (req, res) => {
  try {
    const catalyst = res.locals.catalyst;
    const { action } = req.query;
    const body = req.body;
    
    console.log(`Calendar Bulk API POST Request - Action: ${action}`);
    
    let result;
    
    switch (action) {
      case 'add':
        if (!body) {
          result = {
            status: 'error',
            message: 'Request body is required for add action'
          };
        } else {
          result = await addCalendarBulkData(catalyst, body);
        }
        break;
        
      case 'bulkImport':
        if (!body || !Array.isArray(body)) {
          result = {
            status: 'error',
            message: 'Request body must be an array for bulkImport action'
          };
        } else {
          result = await bulkImportCalendarData(catalyst, body);
        }
        break;
        
      default:
        result = {
          status: 'error',
          message: 'Invalid action. Supported actions: add, bulkImport'
        };
    }
    
    res.status(result.status === 'success' ? 200 : 400).json(result);
    
  } catch (error) {
    console.error('Calendar Bulk API Error:', error);
    res.status(500).json({
      status: 'error',
      message: 'Internal server error',
      error: error.message
    });
  }
});

// PUT route for updates
app.put('/calenderbulk', async (req, res) => {
  try {
    const catalyst = res.locals.catalyst;
    const { action, id } = req.query;
    const body = req.body;
    
    console.log(`Calendar Bulk API PUT Request - Action: ${action}, ID: ${id}`);
    
    let result;
    
    switch (action) {
      case 'update':
        if (!id || !body) {
          result = {
            status: 'error',
            message: 'ID and request body are required for update action'
          };
        } else {
          result = await updateCalendarBulkData(catalyst, id, body);
        }
        break;
        
      default:
        result = {
          status: 'error',
          message: 'Invalid action. Supported actions: update'
        };
    }
    
    res.status(result.status === 'success' ? 200 : 400).json(result);
    
  } catch (error) {
    console.error('Calendar Bulk API Error:', error);
    res.status(500).json({
      status: 'error',
      message: 'Internal server error',
      error: error.message
    });
  }
});

// DELETE route for deletions
app.delete('/calenderbulk', async (req, res) => {
  try {
    const catalyst = res.locals.catalyst;
    const { action, id } = req.query;
    
    console.log(`Calendar Bulk API DELETE Request - Action: ${action}, ID: ${id}`);
    
    let result;
    
    switch (action) {
      case 'delete':
        if (!id) {
          result = {
            status: 'error',
            message: 'ID parameter is required for delete action'
          };
        } else {
          result = await deleteCalendarBulkData(catalyst, id);
        }
        break;
        
      default:
        result = {
          status: 'error',
          message: 'Invalid action. Supported actions: delete'
        };
    }
    
    res.status(result.status === 'success' ? 200 : 400).json(result);
    
  } catch (error) {
    console.error('Calendar Bulk API Error:', error);
    res.status(500).json({
      status: 'error',
      message: 'Internal server error',
      error: error.message
    });
  }
});

module.exports = app;
