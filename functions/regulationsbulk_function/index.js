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

// Table name for regulations bulk data - make sure this matches exactly with your Data Store table name
// From the image, the table appears to be named "RegulationsBulk"
const TABLE_NAME = 'RegulationsBulk';

// Regulations bulk data structure based on the Data Store schema
const createRegulationsBulkRecord = (data) => {
  return {
    Sector: data.sector || '',
    Regulations: data.regulations || '',
    Type: data.type || '',
    States: data.states || '',
    Description: data.description || ''
  };
};

// Convert Data Store record back to application format
const convertToAppFormat = (record) => {
  return {
    id: record.ROWID,
    sector: record.Sector || '',
    regulations: record.Regulations || '',
    type: record.Type || '',
    states: record.States || '',
    description: record.Description || '',
    createdTime: record.CREATEDTIME,
    modifiedTime: record.MODIFIEDTIME,
    creatorId: record.CREATORID
  };
};

// Get all regulations bulk data
const getAllRegulationsBulkData = async (catalyst) => {
  try {
    console.log('Fetching all regulations bulk data from Data Store...');
    
    const dataStore = catalyst.datastore();
    const table = dataStore.table(TABLE_NAME);
    
    // Use getAllRows method instead of select
    const result = await table.getAllRows();
    
    console.log(`Found ${result.length} regulations bulk records`);
    
    const formattedData = result.map(record => convertToAppFormat(record));
    
    return {
      status: 'success',
      message: 'Regulations bulk data retrieved successfully',
      data: formattedData,
      count: formattedData.length
    };
  } catch (error) {
    console.error('Error fetching regulations bulk data:', error);
    return {
      status: 'error',
      message: 'Failed to fetch regulations bulk data',
      error: error.message
    };
  }
};

// Get regulations bulk data by ID
const getRegulationsBulkById = async (catalyst, id) => {
  try {
    console.log(`Fetching regulations bulk data for ID: ${id}`);
    
    const dataStore = catalyst.datastore();
    const table = dataStore.table(TABLE_NAME);
    
    // Use getRow method instead of select
    const result = await table.getRow(id);
    
    if (!result) {
      return {
        status: 'error',
        message: 'Regulations bulk record not found'
      };
    }
    
    const formattedData = convertToAppFormat(result);
    
    return {
      status: 'success',
      message: 'Regulations bulk data retrieved successfully',
      data: formattedData
    };
  } catch (error) {
    console.error('Error fetching regulations bulk data by ID:', error);
    return {
      status: 'error',
      message: 'Failed to fetch regulations bulk data',
      error: error.message
    };
  }
};

// Add new regulations bulk data
const addRegulationsBulkData = async (catalyst, data) => {
  try {
    console.log('Adding new regulations bulk data:', data);
    console.log('Table name being used:', TABLE_NAME);
    
    const dataStore = catalyst.datastore();
    const table = dataStore.table(TABLE_NAME);
    const record = createRegulationsBulkRecord(data);
    
    console.log('Record to be inserted:', record);
    
    const result = await table.insertRow(record);
    
    console.log('Regulations bulk data added successfully with ROWID:', result.ROWID);
    console.log('Full result object:', result);
    
    return {
      status: 'success',
      message: 'Regulations bulk data added successfully',
      data: {
        id: result.ROWID,
        ...data
      }
    };
  } catch (error) {
    console.error('Error adding regulations bulk data:', error);
    console.error('Error details:', error.message);
    console.error('Error stack:', error.stack);
    return {
      status: 'error',
      message: 'Failed to add regulations bulk data',
      error: error.message
    };
  }
};

// Update regulations bulk data
const updateRegulationsBulkData = async (catalyst, id, data) => {
  try {
    console.log(`Updating regulations bulk data for ID: ${id}`, data);
    
    const dataStore = catalyst.datastore();
    const table = dataStore.table(TABLE_NAME);
    const record = createRegulationsBulkRecord(data);
    
    const result = await table.updateRow(id, record);
    
    console.log('Regulations bulk data updated successfully');
    
    return {
      status: 'success',
      message: 'Regulations bulk data updated successfully',
      data: {
        id: id,
        ...data
      }
    };
  } catch (error) {
    console.error('Error updating regulations bulk data:', error);
    return {
      status: 'error',
      message: 'Failed to update regulations bulk data',
      error: error.message
    };
  }
};

// Delete regulations bulk data
const deleteRegulationsBulkData = async (catalyst, id) => {
  try {
    console.log(`Deleting regulations bulk data for ID: ${id}`);
    
    const dataStore = catalyst.datastore();
    const table = dataStore.table(TABLE_NAME);
    await table.deleteRow(id);
    
    console.log('Regulations bulk data deleted successfully');
    
    return {
      status: 'success',
      message: 'Regulations bulk data deleted successfully'
    };
  } catch (error) {
    console.error('Error deleting regulations bulk data:', error);
    return {
      status: 'error',
      message: 'Failed to delete regulations bulk data',
      error: error.message
    };
  }
};

// Bulk import regulations data
const bulkImportRegulationsData = async (catalyst, dataArray) => {
  try {
    console.log(`Bulk importing ${dataArray.length} regulations bulk records`);
    console.log('Table name being used:', TABLE_NAME);
    
    const dataStore = catalyst.datastore();
    const table = dataStore.table(TABLE_NAME);
    const records = dataArray.map(data => createRegulationsBulkRecord(data));
    
    console.log('Records to be inserted:', records.slice(0, 2)); // Log first 2 records as sample
    
    const result = await table.insertRows(records);
    
    console.log(`Bulk import completed. ${result.length} records added`);
    console.log('Result details:', result.slice(0, 2)); // Log first 2 results as sample
    
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
    console.error('Error details:', error.message);
    console.error('Error stack:', error.stack);
    return {
      status: 'error',
      message: 'Failed to bulk import regulations data',
      error: error.message
    };
  }
};

// Get regulations data by sector
const getRegulationsDataBySector = async (catalyst, sector) => {
  try {
    console.log(`Fetching regulations data for sector: ${sector}`);
    
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
      message: `Regulations data for sector '${sector}' retrieved successfully`,
      data: formattedData,
      count: formattedData.length
    };
  } catch (error) {
    console.error('Error fetching regulations data by sector:', error);
    return {
      status: 'error',
      message: 'Failed to fetch regulations data by sector',
      error: error.message
    };
  }
};

// Get regulations data by type
const getRegulationsDataByType = async (catalyst, type) => {
  try {
    console.log(`Fetching regulations data for type: ${type}`);
    
    const dataStore = catalyst.datastore();
    const table = dataStore.table(TABLE_NAME);
    
    // Get all rows and filter by type
    const allRows = await table.getAllRows();
    const filteredRows = allRows.filter(record => 
      record.Type && record.Type.toLowerCase() === type.toLowerCase()
    );
    
    console.log(`Found ${filteredRows.length} records for type: ${type}`);
    
    const formattedData = filteredRows.map(record => convertToAppFormat(record));
    
    return {
      status: 'success',
      message: `Regulations data for type '${type}' retrieved successfully`,
      data: formattedData,
      count: formattedData.length
    };
  } catch (error) {
    console.error('Error fetching regulations data by type:', error);
    return {
      status: 'error',
      message: 'Failed to fetch regulations data by type',
      error: error.message
    };
  }
};

// Get regulations data by states
const getRegulationsDataByStates = async (catalyst, states) => {
  try {
    console.log(`Fetching regulations data for states: ${states}`);
    
    const dataStore = catalyst.datastore();
    const table = dataStore.table(TABLE_NAME);
    
    // Get all rows and filter by states
    const allRows = await table.getAllRows();
    const filteredRows = allRows.filter(record => 
      record.States && record.States.toLowerCase().includes(states.toLowerCase())
    );
    
    console.log(`Found ${filteredRows.length} records for states: ${states}`);
    
    const formattedData = filteredRows.map(record => convertToAppFormat(record));
    
    return {
      status: 'success',
      message: `Regulations data for states '${states}' retrieved successfully`,
      data: formattedData,
      count: formattedData.length
    };
  } catch (error) {
    console.error('Error fetching regulations data by states:', error);
    return {
      status: 'error',
      message: 'Failed to fetch regulations data by states',
      error: error.message
    };
  }
};

// Get count of regulations bulk records
const getRegulationsBulkCount = async (catalyst) => {
  try {
    console.log('Getting regulations bulk data count...');
    
    const dataStore = catalyst.datastore();
    const table = dataStore.table(TABLE_NAME);
    
    // Use getAllRows and get length
    const result = await table.getAllRows();
    
    return {
      status: 'success',
      message: 'Regulations bulk data count retrieved successfully',
      count: result.length
    };
  } catch (error) {
    console.error('Error getting regulations bulk count:', error);
    return {
      status: 'error',
      message: 'Failed to get regulations bulk data count',
      error: error.message
    };
  }
};

// Health check
app.get('/', (req, res) => {
  res.status(200).json({ status: 'success', message: 'regulationsbulk_function ready' });
});

// Test endpoint to check table access and verify data
app.get('/test-table', async (req, res) => {
  try {
    console.log('Testing RegulationsBulk table access...');
    const catalyst = res.locals.catalyst;
    const dataStore = catalyst.datastore();
    const table = dataStore.table(TABLE_NAME);
    
    console.log('Table name:', TABLE_NAME);
    console.log('Table object:', table);
    
    // Try to get all rows to see if table exists and has data
    const allRows = await table.getAllRows();
    console.log('Total rows in table:', allRows.length);
    console.log('Sample rows:', allRows.slice(0, 3));
    
    // Also try ZCQL query as alternative
    const zcql = catalyst.zcql();
    const countQuery = await zcql.executeZCQLQuery(`SELECT COUNT(ROWID) as count FROM ${TABLE_NAME}`);
    console.log('ZCQL count result:', countQuery);
    
    res.status(200).json({ 
      status: 'success', 
      message: 'Table access successful',
      tableInfo: {
        tableName: TABLE_NAME,
        recordCount: allRows.length,
        zcqlCount: countQuery[0]?.[TABLE_NAME]?.count || 0,
        sampleData: allRows.slice(0, 3)
      }
    });
  } catch (err) {
    console.error('Table test error:', err);
    res.status(500).json({ 
      status: 'failure', 
      message: err.message || 'Failed to access table',
      error: err.toString(),
      tableName: TABLE_NAME
    });
  }
});

// Test endpoint to add a single record for debugging
app.post('/test-add', async (req, res) => {
  try {
    console.log('Test add endpoint called');
    const catalyst = res.locals.catalyst;
    
    const testData = {
      sector: "Test Sector",
      regulations: "Test Regulation",
      type: "Test Type",
      states: "Test State",
      description: "Test Description"
    };
    
    console.log('Test data to be added:', testData);
    
    const result = await addRegulationsBulkData(catalyst, testData);
    
    res.status(result.status === 'success' ? 200 : 400).json(result);
    
  } catch (error) {
    console.error('Test add error:', error);
    res.status(500).json({
      status: 'error',
      message: 'Test add failed',
      error: error.message
    });
  }
});

// Get all regulations bulk data
app.get('/regulationsbulk', async (req, res) => {
  try {
    const catalyst = res.locals.catalyst;
    const { action, id, sector, type, states } = req.query;
    
    console.log(`Regulations Bulk API Request - Action: ${action}, ID: ${id}, Sector: ${sector}, Type: ${type}, States: ${states}`);
    
    let result;
    
    switch (action) {
      case 'getAll':
        result = await getAllRegulationsBulkData(catalyst);
        break;
        
      case 'getById':
        if (!id) {
          result = {
            status: 'error',
            message: 'ID parameter is required for getById action'
          };
        } else {
          result = await getRegulationsBulkById(catalyst, id);
        }
        break;
        
      case 'getBySector':
        if (!sector) {
          result = {
            status: 'error',
            message: 'Sector parameter is required for getBySector action'
          };
        } else {
          result = await getRegulationsDataBySector(catalyst, sector);
        }
        break;
        
      case 'getByType':
        if (!type) {
          result = {
            status: 'error',
            message: 'Type parameter is required for getByType action'
          };
        } else {
          result = await getRegulationsDataByType(catalyst, type);
        }
        break;
        
      case 'getByStates':
        if (!states) {
          result = {
            status: 'error',
            message: 'States parameter is required for getByStates action'
          };
        } else {
          result = await getRegulationsDataByStates(catalyst, states);
        }
        break;
        
      case 'count':
        result = await getRegulationsBulkCount(catalyst);
        break;
        
      default:
        result = {
          status: 'error',
          message: 'Invalid action. Supported actions: getAll, getById, getBySector, getByType, getByStates, count'
        };
    }
    
    res.status(result.status === 'success' ? 200 : 400).json(result);
    
  } catch (error) {
    console.error('Regulations Bulk API Error:', error);
    res.status(500).json({
      status: 'error',
      message: 'Internal server error',
      error: error.message
    });
  }
});

// POST routes for data modification
app.post('/regulationsbulk', async (req, res) => {
  try {
    const catalyst = res.locals.catalyst;
    const { action } = req.query;
    const body = req.body;
    
    console.log(`Regulations Bulk API POST Request - Action: ${action}`);
    
    let result;
    
    switch (action) {
      case 'add':
        if (!body) {
          result = {
            status: 'error',
            message: 'Request body is required for add action'
          };
        } else {
          result = await addRegulationsBulkData(catalyst, body);
        }
        break;
        
      case 'bulkImport':
        if (!body || !Array.isArray(body)) {
          result = {
            status: 'error',
            message: 'Request body must be an array for bulkImport action'
          };
        } else {
          result = await bulkImportRegulationsData(catalyst, body);
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
    console.error('Regulations Bulk API Error:', error);
    res.status(500).json({
      status: 'error',
      message: 'Internal server error',
      error: error.message
    });
  }
});

// PUT route for updates
app.put('/regulationsbulk', async (req, res) => {
  try {
    const catalyst = res.locals.catalyst;
    const { action, id } = req.query;
    const body = req.body;
    
    console.log(`Regulations Bulk API PUT Request - Action: ${action}, ID: ${id}`);
    
    let result;
    
    switch (action) {
      case 'update':
        if (!id || !body) {
          result = {
            status: 'error',
            message: 'ID and request body are required for update action'
          };
        } else {
          result = await updateRegulationsBulkData(catalyst, id, body);
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
    console.error('Regulations Bulk API Error:', error);
    res.status(500).json({
      status: 'error',
      message: 'Internal server error',
      error: error.message
    });
  }
});

// DELETE route for deletions
app.delete('/regulationsbulk', async (req, res) => {
  try {
    const catalyst = res.locals.catalyst;
    const { action, id } = req.query;
    
    console.log(`Regulations Bulk API DELETE Request - Action: ${action}, ID: ${id}`);
    
    let result;
    
    switch (action) {
      case 'delete':
        if (!id) {
          result = {
            status: 'error',
            message: 'ID parameter is required for delete action'
          };
        } else {
          result = await deleteRegulationsBulkData(catalyst, id);
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
    console.error('Regulations Bulk API Error:', error);
    res.status(500).json({
      status: 'error',
      message: 'Internal server error',
      error: error.message
    });
  }
});

module.exports = app;