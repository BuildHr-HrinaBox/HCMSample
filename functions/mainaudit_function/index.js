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

// Table name for main audit data - make sure this matches exactly with your Data Store table name
const TABLE_NAME = 'MainAudit';

// Main audit data structure based on the Data Store schema
// Column order: Act, Description, RecordCategory, MaximumMarks, Applicability, Sector, Remarks
const createMainAuditRecord = (data) => {
  return {
    Act: data.act || '',
    Description: data.description || '',
    RecordCategory: data.recordCategory || '',
    MaximumMarks: data.maximumMarks || '',
    Applicability: data.applicability || '',
    Sector: data.sector || '',
    Remarks: data.remarks || ''
  };
};

// Convert Data Store record back to application format
const convertToAppFormat = (record) => {
  return {
    id: record.ROWID,
    act: record.Act || '',
    description: record.Description || '',
    recordCategory: record.RecordCategory || '',
    maximumMarks: record.MaximumMarks || '',
    applicability: record.Applicability || '',
    sector: record.Sector || '',
    remarks: record.Remarks || '',
    createdTime: record.CREATEDTIME,
    modifiedTime: record.MODIFIEDTIME,
    creatorId: record.CREATORID
  };
};

// Get all main audit data
const getAllMainAuditData = async (catalyst) => {
  try {
    console.log('Fetching all main audit data from Data Store...');
    
    const dataStore = catalyst.datastore();
    const table = dataStore.table(TABLE_NAME);
    
    const result = await table.getAllRows();
    
    console.log(`Found ${result.length} main audit records`);
    
    const formattedData = result.map(record => convertToAppFormat(record));
    
    return {
      status: 'success',
      message: 'Main audit data retrieved successfully',
      data: formattedData,
      count: formattedData.length
    };
  } catch (error) {
    console.error('Error fetching main audit data:', error);
    return {
      status: 'error',
      message: 'Failed to fetch main audit data',
      error: error.message
    };
  }
};

// Get main audit data by ID
const getMainAuditById = async (catalyst, id) => {
  try {
    console.log(`Fetching main audit data for ID: ${id}`);
    
    const dataStore = catalyst.datastore();
    const table = dataStore.table(TABLE_NAME);
    
    const result = await table.getRow(id);
    
    if (!result) {
      return {
        status: 'error',
        message: 'Main audit record not found'
      };
    }
    
    const formattedData = convertToAppFormat(result);
    
    return {
      status: 'success',
      message: 'Main audit data retrieved successfully',
      data: formattedData
    };
  } catch (error) {
    console.error('Error fetching main audit data by ID:', error);
    return {
      status: 'error',
      message: 'Failed to fetch main audit data',
      error: error.message
    };
  }
};

// Add new main audit data
const addMainAuditData = async (catalyst, data) => {
  try {
    console.log('Adding new main audit data:', data);
    console.log('Table name being used:', TABLE_NAME);
    
    // Validate required field
    if (!data.act || !data.act.trim()) {
      return {
        status: 'error',
        message: 'Act field is required'
      };
    }
    
    const dataStore = catalyst.datastore();
    const table = dataStore.table(TABLE_NAME);
    const record = createMainAuditRecord(data);
    
    console.log('Record to be inserted:', record);
    
    const result = await table.insertRow(record);
    
    console.log('Main audit data added successfully with ROWID:', result.ROWID);
    
    return {
      status: 'success',
      message: 'Main audit data added successfully',
      data: {
        id: result.ROWID,
        ...data
      }
    };
  } catch (error) {
    console.error('Error adding main audit data:', error);
    console.error('Error details:', error.message);
    console.error('Error stack:', error.stack);
    return {
      status: 'error',
      message: 'Failed to add main audit data',
      error: error.message
    };
  }
};

// Update main audit data
const updateMainAuditData = async (catalyst, id, data) => {
  try {
    console.log(`Updating main audit data for ID: ${id}`, data);
    
    const dataStore = catalyst.datastore();
    const table = dataStore.table(TABLE_NAME);
    const record = createMainAuditRecord(data);
    
    const result = await table.updateRow(id, record);
    
    console.log('Main audit data updated successfully');
    
    return {
      status: 'success',
      message: 'Main audit data updated successfully',
      data: {
        id: id,
        ...data
      }
    };
  } catch (error) {
    console.error('Error updating main audit data:', error);
    return {
      status: 'error',
      message: 'Failed to update main audit data',
      error: error.message
    };
  }
};

// Delete main audit data
const deleteMainAuditData = async (catalyst, id) => {
  try {
    console.log(`Deleting main audit data for ID: ${id}`);
    
    const dataStore = catalyst.datastore();
    const table = dataStore.table(TABLE_NAME);
    await table.deleteRow(id);
    
    console.log('Main audit data deleted successfully');
    
    return {
      status: 'success',
      message: 'Main audit data deleted successfully'
    };
  } catch (error) {
    console.error('Error deleting main audit data:', error);
    return {
      status: 'error',
      message: 'Failed to delete main audit data',
      error: error.message
    };
  }
};

// Delete all main audit data (bulk delete)
const deleteAllMainAuditData = async (catalyst) => {
  try {
    console.log('Deleting all main audit data from Data Store...');
    
    const dataStore = catalyst.datastore();
    const table = dataStore.table(TABLE_NAME);
    
    const allRows = await table.getAllRows();
    console.log(`Found ${allRows.length} records to delete`);
    
    if (allRows.length === 0) {
      return {
        status: 'success',
        message: 'No records found to delete',
        deletedCount: 0
      };
    }
    
    let deletedCount = 0;
    let failedCount = 0;
    const errors = [];
    
    for (let i = 0; i < allRows.length; i++) {
      try {
        await table.deleteRow(allRows[i].ROWID);
        deletedCount++;
        
        if ((i + 1) % 10 === 0) {
          console.log(`Deleted ${i + 1}/${allRows.length} records...`);
        }
      } catch (deleteError) {
        failedCount++;
        errors.push({
          rowId: allRows[i].ROWID,
          error: deleteError.message
        });
        console.error(`Failed to delete row ${allRows[i].ROWID}:`, deleteError.message);
      }
    }
    
    console.log(`Deletion completed. Successfully deleted: ${deletedCount}, Failed: ${failedCount}`);
    
    if (failedCount === 0) {
      return {
        status: 'success',
        message: `Successfully deleted ${deletedCount} main audit records`,
        deletedCount: deletedCount
      };
    } else if (deletedCount > 0) {
      return {
        status: 'partial',
        message: `Partially deleted: ${deletedCount} succeeded, ${failedCount} failed`,
        deletedCount: deletedCount,
        failedCount: failedCount,
        errors: errors
      };
    } else {
      return {
        status: 'error',
        message: `Failed to delete all records. ${failedCount} records failed to delete`,
        deletedCount: 0,
        failedCount: failedCount,
        errors: errors
      };
    }
  } catch (error) {
    console.error('Error deleting all main audit data:', error);
    console.error('Error stack:', error.stack);
    return {
      status: 'error',
      message: 'Failed to delete all main audit data',
      error: error.message || error.toString()
    };
  }
};

// Bulk import main audit data
const bulkImportMainAuditData = async (catalyst, dataArray) => {
  try {
    console.log(`Bulk importing ${dataArray.length} main audit records`);
    console.log('Table name being used:', TABLE_NAME);
    
    // Validate that all records have required Act field
    const invalidRecords = dataArray.filter((data, index) => !data.act || !data.act.trim());
    if (invalidRecords.length > 0) {
      return {
        status: 'error',
        message: `${invalidRecords.length} record(s) are missing the required Act field`,
        invalidRecords: invalidRecords.length
      };
    }
    
    const dataStore = catalyst.datastore();
    const table = dataStore.table(TABLE_NAME);
    const records = dataArray.map(data => createMainAuditRecord(data));
    
    console.log('Records to be inserted:', records.slice(0, 2));
    
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
    console.error('Error details:', error.message);
    console.error('Error stack:', error.stack);
    return {
      status: 'error',
      message: 'Failed to bulk import main audit data',
      error: error.message
    };
  }
};

// Get main audit data by Act
const getMainAuditByAct = async (catalyst, act) => {
  try {
    console.log(`Fetching main audit data for Act: ${act}`);
    
    const dataStore = catalyst.datastore();
    const table = dataStore.table(TABLE_NAME);
    
    const allRows = await table.getAllRows();
    const filteredRows = allRows.filter(record => 
      record.Act && record.Act.toLowerCase().includes(act.toLowerCase())
    );
    
    console.log(`Found ${filteredRows.length} records for Act: ${act}`);
    
    const formattedData = filteredRows.map(record => convertToAppFormat(record));
    
    return {
      status: 'success',
      message: `Main audit data for Act '${act}' retrieved successfully`,
      data: formattedData,
      count: formattedData.length
    };
  } catch (error) {
    console.error('Error fetching main audit data by Act:', error);
    return {
      status: 'error',
      message: 'Failed to fetch main audit data by Act',
      error: error.message
    };
  }
};

// Get main audit data by Sector
const getMainAuditBySector = async (catalyst, sector) => {
  try {
    console.log(`Fetching main audit data for Sector: ${sector}`);
    
    const dataStore = catalyst.datastore();
    const table = dataStore.table(TABLE_NAME);
    
    const allRows = await table.getAllRows();
    const filteredRows = allRows.filter(record => 
      record.Sector && record.Sector.toLowerCase().includes(sector.toLowerCase())
    );
    
    console.log(`Found ${filteredRows.length} records for Sector: ${sector}`);
    
    const formattedData = filteredRows.map(record => convertToAppFormat(record));
    
    return {
      status: 'success',
      message: `Main audit data for Sector '${sector}' retrieved successfully`,
      data: formattedData,
      count: formattedData.length
    };
  } catch (error) {
    console.error('Error fetching main audit data by Sector:', error);
    return {
      status: 'error',
      message: 'Failed to fetch main audit data by Sector',
      error: error.message
    };
  }
};

// Get main audit data by RecordCategory
const getMainAuditByRecordCategory = async (catalyst, recordCategory) => {
  try {
    console.log(`Fetching main audit data for RecordCategory: ${recordCategory}`);
    
    const dataStore = catalyst.datastore();
    const table = dataStore.table(TABLE_NAME);
    
    const allRows = await table.getAllRows();
    const filteredRows = allRows.filter(record => 
      record.RecordCategory && record.RecordCategory.toLowerCase() === recordCategory.toLowerCase()
    );
    
    console.log(`Found ${filteredRows.length} records for RecordCategory: ${recordCategory}`);
    
    const formattedData = filteredRows.map(record => convertToAppFormat(record));
    
    return {
      status: 'success',
      message: `Main audit data for RecordCategory '${recordCategory}' retrieved successfully`,
      data: formattedData,
      count: formattedData.length
    };
  } catch (error) {
    console.error('Error fetching main audit data by RecordCategory:', error);
    return {
      status: 'error',
      message: 'Failed to fetch main audit data by RecordCategory',
      error: error.message
    };
  }
};

// Get count of main audit records
const getMainAuditCount = async (catalyst) => {
  try {
    console.log('Getting main audit data count...');
    
    const dataStore = catalyst.datastore();
    const table = dataStore.table(TABLE_NAME);
    
    const result = await table.getAllRows();
    
    return {
      status: 'success',
      message: 'Main audit data count retrieved successfully',
      count: result.length
    };
  } catch (error) {
    console.error('Error getting main audit count:', error);
    return {
      status: 'error',
      message: 'Failed to get main audit data count',
      error: error.message
    };
  }
};

// Health check
app.get('/', (req, res) => {
  res.status(200).json({ status: 'success', message: 'mainaudit_function ready' });
});

// Test endpoint to check table access
app.get('/test-table', async (req, res) => {
  try {
    console.log('Testing MainAudit table access...');
    const catalyst = res.locals.catalyst;
    const dataStore = catalyst.datastore();
    const table = dataStore.table(TABLE_NAME);
    
    console.log('Table name:', TABLE_NAME);
    
    const allRows = await table.getAllRows();
    console.log('Total rows in table:', allRows.length);
    
    const zcql = catalyst.zcql();
    const countQuery = await zcql.executeZCQLQuery(`SELECT COUNT(ROWID) as count FROM ${TABLE_NAME}`);
    
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

// Test endpoint to add a single record
app.post('/test-add', async (req, res) => {
  try {
    console.log('Test add endpoint called');
    const catalyst = res.locals.catalyst;
    
    const testData = {
      act: "Test Act",
      description: "Test Description",
      recordCategory: "Test Category",
      maximumMarks: "100",
      applicability: "Test Applicability",
      sector: "Test Sector"
    };
    
    console.log('Test data to be added:', testData);
    
    const result = await addMainAuditData(catalyst, testData);
    
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

// Get all main audit data
app.get('/mainaudit', async (req, res) => {
  try {
    const catalyst = res.locals.catalyst;
    const { action, id, act, sector, recordCategory } = req.query;
    
    console.log(`Main Audit API Request - Action: ${action}, ID: ${id}, Act: ${act}, Sector: ${sector}, RecordCategory: ${recordCategory}`);
    
    let result;
    
    switch (action) {
      case 'getAll':
        result = await getAllMainAuditData(catalyst);
        break;
        
      case 'getById':
        if (!id) {
          result = {
            status: 'error',
            message: 'ID parameter is required for getById action'
          };
        } else {
          result = await getMainAuditById(catalyst, id);
        }
        break;
        
      case 'getByAct':
        if (!act) {
          result = {
            status: 'error',
            message: 'Act parameter is required for getByAct action'
          };
        } else {
          result = await getMainAuditByAct(catalyst, act);
        }
        break;
        
      case 'getBySector':
        if (!sector) {
          result = {
            status: 'error',
            message: 'Sector parameter is required for getBySector action'
          };
        } else {
          result = await getMainAuditBySector(catalyst, sector);
        }
        break;
        
      case 'getByRecordCategory':
        if (!recordCategory) {
          result = {
            status: 'error',
            message: 'RecordCategory parameter is required for getByRecordCategory action'
          };
        } else {
          result = await getMainAuditByRecordCategory(catalyst, recordCategory);
        }
        break;
        
      case 'count':
        result = await getMainAuditCount(catalyst);
        break;
        
      default:
        result = {
          status: 'error',
          message: 'Invalid action. Supported actions: getAll, getById, getByAct, getBySector, getByRecordCategory, count'
        };
    }
    
    res.status(result.status === 'success' ? 200 : 400).json(result);
    
  } catch (error) {
    console.error('Main Audit API Error:', error);
    res.status(500).json({
      status: 'error',
      message: 'Internal server error',
      error: error.message
    });
  }
});

// POST routes for data modification
app.post('/mainaudit', async (req, res) => {
  try {
    const catalyst = res.locals.catalyst;
    const { action } = req.query;
    const body = req.body;
    
    console.log(`Main Audit API POST Request - Action: ${action}`);
    
    let result;
    
    switch (action) {
      case 'add':
        if (!body) {
          result = {
            status: 'error',
            message: 'Request body is required for add action'
          };
        } else {
          result = await addMainAuditData(catalyst, body);
        }
        break;
        
      case 'bulkImport':
        if (!body || !Array.isArray(body)) {
          result = {
            status: 'error',
            message: 'Request body must be an array for bulkImport action'
          };
        } else {
          result = await bulkImportMainAuditData(catalyst, body);
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
    console.error('Main Audit API Error:', error);
    res.status(500).json({
      status: 'error',
      message: 'Internal server error',
      error: error.message
    });
  }
});

// PUT route for updates
app.put('/mainaudit', async (req, res) => {
  try {
    const catalyst = res.locals.catalyst;
    const { action, id } = req.query;
    const body = req.body;
    
    console.log(`Main Audit API PUT Request - Action: ${action}, ID: ${id}`);
    
    let result;
    
    switch (action) {
      case 'update':
        if (!id || !body) {
          result = {
            status: 'error',
            message: 'ID and request body are required for update action'
          };
        } else {
          result = await updateMainAuditData(catalyst, id, body);
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
    console.error('Main Audit API Error:', error);
    res.status(500).json({
      status: 'error',
      message: 'Internal server error',
      error: error.message
    });
  }
});

// DELETE route for deletions
app.delete('/mainaudit', async (req, res) => {
  try {
    const catalyst = res.locals.catalyst;
    const { action, id } = req.query;
    
    console.log(`Main Audit API DELETE Request - Action: ${action}, ID: ${id}`);
    
    let result;
    
    switch (action) {
      case 'delete':
        if (!id) {
          result = {
            status: 'error',
            message: 'ID parameter is required for delete action'
          };
        } else {
          result = await deleteMainAuditData(catalyst, id);
        }
        break;
        
      case 'deleteAll':
        result = await deleteAllMainAuditData(catalyst);
        break;
        
      default:
        result = {
          status: 'error',
          message: 'Invalid action. Supported actions: delete, deleteAll'
        };
    }
    
    res.status(result.status === 'success' ? 200 : 400).json(result);
    
  } catch (error) {
    console.error('Main Audit API Error:', error);
    res.status(500).json({
      status: 'error',
      message: 'Internal server error',
      error: error.message
    });
  }
});

module.exports = app;
