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

// Table name for audit master data - make sure this matches exactly with your Data Store table name
const TABLE_NAME = 'Audit';

// Audit master data structure based on the Data Store schema
const createAuditMasterRecord = (data) => {
  return {
    Sector: data.sector || '',
    Act: data.act || '',
    Description: data.description || '',
    RecordCategory: data.recordCategory || '',
    MaximumMarks: data.maximumMarks || '',
    Applicability: data.applicability || '',
    Remarks: data.remarks || '',
    SentForApproval: data.sentForApproval || 'false',
    ApprovedForSite: data.approvedForSite || ''
  };
};

// Convert Data Store record back to application format
const convertToAppFormat = (record) => {
  return {
    id: record.ROWID,
    sector: record.Sector || '',
    act: record.Act || '',
    description: record.Description || '',
    recordCategory: record.RecordCategory || '',
    maximumMarks: record.MaximumMarks || '',
    applicability: record.Applicability || '',
    remarks: record.Remarks || '',
    sentForApproval: record.SentForApproval || 'false',
    approvedForSite: record.ApprovedForSite || '',
    createdTime: record.CREATEDTIME,
    modifiedTime: record.MODIFIEDTIME,
    creatorId: record.CREATORID
  };
};

// Get all audit master data
const getAllAuditMasterData = async (catalyst) => {
  try {
    console.log('Fetching all audit master data from Data Store...');
    
    const dataStore = catalyst.datastore();
    const table = dataStore.table(TABLE_NAME);
    
    // Use getAllRows method instead of select
    const result = await table.getAllRows();
    
    console.log(`Found ${result.length} audit master records`);
    
    const formattedData = result.map(record => convertToAppFormat(record));
    
    return {
      status: 'success',
      message: 'Audit master data retrieved successfully',
      data: formattedData,
      count: formattedData.length
    };
  } catch (error) {
    console.error('Error fetching audit master data:', error);
    return {
      status: 'error',
      message: 'Failed to fetch audit master data',
      error: error.message
    };
  }
};

// Get audit master data by ID
const getAuditMasterById = async (catalyst, id) => {
  try {
    console.log(`Fetching audit master data for ID: ${id}`);
    
    const dataStore = catalyst.datastore();
    const table = dataStore.table(TABLE_NAME);
    
    // Use getRow method instead of select
    const result = await table.getRow(id);
    
    if (!result) {
      return {
        status: 'error',
        message: 'Audit master record not found'
      };
    }
    
    const formattedData = convertToAppFormat(result);
    
    return {
      status: 'success',
      message: 'Audit master data retrieved successfully',
      data: formattedData
    };
  } catch (error) {
    console.error('Error fetching audit master data by ID:', error);
    return {
      status: 'error',
      message: 'Failed to fetch audit master data',
      error: error.message
    };
  }
};

// Add new audit master data
const addAuditMasterData = async (catalyst, data) => {
  try {
    console.log('Adding new audit master data:', data);
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
    const record = createAuditMasterRecord(data);
    
    console.log('Record to be inserted:', record);
    
    const result = await table.insertRow(record);
    
    console.log('Audit master data added successfully with ROWID:', result.ROWID);
    console.log('Full result object:', result);
    
    return {
      status: 'success',
      message: 'Audit master data added successfully',
      data: {
        id: result.ROWID,
        ...data
      }
    };
  } catch (error) {
    console.error('Error adding audit master data:', error);
    console.error('Error details:', error.message);
    console.error('Error stack:', error.stack);
    return {
      status: 'error',
      message: 'Failed to add audit master data',
      error: error.message
    };
  }
};

// Update audit master data
const updateAuditMasterData = async (catalyst, id, data) => {
  try {
    console.log(`Updating audit master data for ID: ${id}`, data);
    console.log(`ID type: ${typeof id}, ID value: ${id}`);
    
    if (!id) {
      throw new Error('ID is required for update operation');
    }
    
    const dataStore = catalyst.datastore();
    const table = dataStore.table(TABLE_NAME);
    
    // First, verify the row exists and get existing data
    let existingRow;
    try {
      existingRow = await table.getRow(id);
      if (!existingRow) {
        throw new Error(`Record with ID ${id} not found`);
      }
      console.log('Existing row found:', existingRow);
    } catch (getError) {
      console.error('Error getting existing row:', getError);
      throw new Error(`Record with ID ${id} not found: ${getError.message}`);
    }
    
    // Preserve SentForApproval and ApprovedForSite if not provided in update data
    const updateData = {
      ...data,
      sentForApproval: data.sentForApproval !== undefined ? data.sentForApproval : (existingRow.SentForApproval || 'false'),
      approvedForSite: data.approvedForSite !== undefined ? data.approvedForSite : (existingRow.ApprovedForSite || '')
    };
    
    const record = createAuditMasterRecord(updateData);
    // Include ROWID in the record object for updateRow
    record.ROWID = id;
    console.log('Record to update (with ROWID):', record);
    
    // updateRow expects a single object with ROWID included
    const result = await table.updateRow(record);
    console.log('Update result:', result);
    
    console.log('Audit master data updated successfully');
    
    return {
      status: 'success',
      message: 'Audit master data updated successfully',
      data: {
        id: id,
        ...data
      }
    };
  } catch (error) {
    console.error('Error updating audit master data:', error);
    console.error('Error stack:', error.stack);
    return {
      status: 'error',
      message: `Failed to update audit master data: ${error.message || error}`,
      error: error.message || String(error)
    };
  }
};

// Delete audit master data
const deleteAuditMasterData = async (catalyst, id) => {
  try {
    console.log(`Deleting audit master data for ID: ${id}`);
    
    const dataStore = catalyst.datastore();
    const table = dataStore.table(TABLE_NAME);
    await table.deleteRow(id);
    
    console.log('Audit master data deleted successfully');
    
    return {
      status: 'success',
      message: 'Audit master data deleted successfully'
    };
  } catch (error) {
    console.error('Error deleting audit master data:', error);
    return {
      status: 'error',
      message: 'Failed to delete audit master data',
      error: error.message
    };
  }
};

// Delete all audit master data (bulk delete)
const deleteAllAuditMasterData = async (catalyst) => {
  try {
    console.log('Deleting all audit master data from Data Store...');
    
    const dataStore = catalyst.datastore();
    const table = dataStore.table(TABLE_NAME);
    
    // Get all rows first
    const allRows = await table.getAllRows();
    console.log(`Found ${allRows.length} records to delete`);
    
    if (allRows.length === 0) {
      return {
        status: 'success',
        message: 'No records found to delete',
        deletedCount: 0
      };
    }
    
    // Delete rows sequentially to avoid overwhelming the API
    let deletedCount = 0;
    let failedCount = 0;
    const errors = [];
    
    for (let i = 0; i < allRows.length; i++) {
      try {
        await table.deleteRow(allRows[i].ROWID);
        deletedCount++;
        
        // Log progress for every 10 records
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
        // Continue with next record even if one fails
      }
    }
    
    console.log(`Deletion completed. Successfully deleted: ${deletedCount}, Failed: ${failedCount}`);
    
    if (failedCount === 0) {
      return {
        status: 'success',
        message: `Successfully deleted ${deletedCount} audit master records`,
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
    console.error('Error deleting all audit master data:', error);
    console.error('Error stack:', error.stack);
    return {
      status: 'error',
      message: 'Failed to delete all audit master data',
      error: error.message || error.toString()
    };
  }
};

// Bulk import audit master data
const bulkImportAuditMasterData = async (catalyst, dataArray) => {
  try {
    console.log(`Bulk importing ${dataArray.length} audit master records`);
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
    const records = dataArray.map(data => createAuditMasterRecord(data));
    
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
      message: 'Failed to bulk import audit master data',
      error: error.message
    };
  }
};

// Get audit master data by Act
const getAuditMasterByAct = async (catalyst, act) => {
  try {
    console.log(`Fetching audit master data for Act: ${act}`);
    
    const dataStore = catalyst.datastore();
    const table = dataStore.table(TABLE_NAME);
    
    // Get all rows and filter by Act
    const allRows = await table.getAllRows();
    const filteredRows = allRows.filter(record => 
      record.Act && record.Act.toLowerCase().includes(act.toLowerCase())
    );
    
    console.log(`Found ${filteredRows.length} records for Act: ${act}`);
    
    const formattedData = filteredRows.map(record => convertToAppFormat(record));
    
    return {
      status: 'success',
      message: `Audit master data for Act '${act}' retrieved successfully`,
      data: formattedData,
      count: formattedData.length
    };
  } catch (error) {
    console.error('Error fetching audit master data by Act:', error);
    return {
      status: 'error',
      message: 'Failed to fetch audit master data by Act',
      error: error.message
    };
  }
};

// Get audit master data by Sector
const getAuditMasterBySector = async (catalyst, sector) => {
  try {
    console.log(`Fetching audit master data for Sector: ${sector}`);
    
    const dataStore = catalyst.datastore();
    const table = dataStore.table(TABLE_NAME);
    
    // Get all rows and filter by Sector
    const allRows = await table.getAllRows();
    const filteredRows = allRows.filter(record => 
      record.Sector && record.Sector.toLowerCase().includes(sector.toLowerCase())
    );
    
    console.log(`Found ${filteredRows.length} records for Sector: ${sector}`);
    
    const formattedData = filteredRows.map(record => convertToAppFormat(record));
    
    return {
      status: 'success',
      message: `Audit master data for Sector '${sector}' retrieved successfully`,
      data: formattedData,
      count: formattedData.length
    };
  } catch (error) {
    console.error('Error fetching audit master data by Sector:', error);
    return {
      status: 'error',
      message: 'Failed to fetch audit master data by Sector',
      error: error.message
    };
  }
};

// Get audit master data by RecordCategory
const getAuditMasterByRecordCategory = async (catalyst, recordCategory) => {
  try {
    console.log(`Fetching audit master data for RecordCategory: ${recordCategory}`);
    
    const dataStore = catalyst.datastore();
    const table = dataStore.table(TABLE_NAME);
    
    // Get all rows and filter by RecordCategory
    const allRows = await table.getAllRows();
    const filteredRows = allRows.filter(record => 
      record.RecordCategory && record.RecordCategory.toLowerCase() === recordCategory.toLowerCase()
    );
    
    console.log(`Found ${filteredRows.length} records for RecordCategory: ${recordCategory}`);
    
    const formattedData = filteredRows.map(record => convertToAppFormat(record));
    
    return {
      status: 'success',
      message: `Audit master data for RecordCategory '${recordCategory}' retrieved successfully`,
      data: formattedData,
      count: formattedData.length
    };
  } catch (error) {
    console.error('Error fetching audit master data by RecordCategory:', error);
    return {
      status: 'error',
      message: 'Failed to fetch audit master data by RecordCategory',
      error: error.message
    };
  }
};

// Get count of audit master records
const getAuditMasterCount = async (catalyst) => {
  try {
    console.log('Getting audit master data count...');
    
    const dataStore = catalyst.datastore();
    const table = dataStore.table(TABLE_NAME);
    
    // Use getAllRows and get length
    const result = await table.getAllRows();
    
    return {
      status: 'success',
      message: 'Audit master data count retrieved successfully',
      count: result.length
    };
  } catch (error) {
    console.error('Error getting audit master count:', error);
    return {
      status: 'error',
      message: 'Failed to get audit master data count',
      error: error.message
    };
  }
};

// Send items for approval (bulk update SentForApproval to 'true' and ApprovedForSite)
const sendItemsForApproval = async (catalyst, itemIds, siteName = '') => {
  try {
    console.log(`Sending ${itemIds.length} items for approval${siteName ? ` for site: ${siteName}` : ''}`);
    
    if (!itemIds || !Array.isArray(itemIds) || itemIds.length === 0) {
      return {
        status: 'error',
        message: 'Item IDs array is required and must not be empty'
      };
    }
    
    const dataStore = catalyst.datastore();
    const table = dataStore.table(TABLE_NAME);
    
    let successCount = 0;
    let failedCount = 0;
    const errors = [];
    
    // Update each item sequentially
    for (let i = 0; i < itemIds.length; i++) {
      const itemId = String(itemIds[i]).trim();
      
      try {
        // Get existing row
        const existingRow = await table.getRow(itemId);
        if (!existingRow) {
          failedCount++;
          errors.push({
            rowId: itemId,
            error: 'Record not found'
          });
          continue;
        }
        
        // Update the record with SentForApproval = 'true' and ApprovedForSite
        const updatedRecord = {
          ROWID: itemId,
          Sector: existingRow.Sector || '',
          Act: existingRow.Act || '',
          Description: existingRow.Description || '',
          RecordCategory: existingRow.RecordCategory || '',
          MaximumMarks: existingRow.MaximumMarks || '',
          Applicability: existingRow.Applicability || '',
          Remarks: existingRow.Remarks || '',
          SentForApproval: 'true',
          ApprovedForSite: siteName || ''
        };
        
        await table.updateRow(updatedRecord);
        successCount++;
        
        // Log progress for every 10 records
        if ((i + 1) % 10 === 0) {
          console.log(`Updated ${i + 1}/${itemIds.length} records...`);
        }
      } catch (updateError) {
        failedCount++;
        errors.push({
          rowId: itemId,
          error: updateError.message
        });
        console.error(`Failed to update row ${itemId}:`, updateError.message);
      }
    }
    
    console.log(`Send for approval completed. Successfully updated: ${successCount}, Failed: ${failedCount}`);
    
    if (failedCount === 0) {
      return {
        status: 'success',
        message: `Successfully sent ${successCount} item(s) for approval`,
        updatedCount: successCount
      };
    } else if (successCount > 0) {
      return {
        status: 'partial',
        message: `Partially sent for approval: ${successCount} succeeded, ${failedCount} failed`,
        updatedCount: successCount,
        failedCount: failedCount,
        errors: errors
      };
    } else {
      return {
        status: 'error',
        message: `Failed to send items for approval. ${failedCount} items failed`,
        updatedCount: 0,
        failedCount: failedCount,
        errors: errors
      };
    }
  } catch (error) {
    console.error('Error sending items for approval:', error);
    return {
      status: 'error',
      message: 'Failed to send items for approval',
      error: error.message
    };
  }
};

// Health check
app.get('/', (req, res) => {
  res.status(200).json({ status: 'success', message: 'auditmaster_function ready' });
});

// Test endpoint to check table access and verify data
app.get('/test-table', async (req, res) => {
  try {
    console.log('Testing Audit table access...');
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
      act: "Test Act",
      description: "Test Description",
      recordCategory: "Test Category",
      maximumMarks: "100",
    };
    
    console.log('Test data to be added:', testData);
    
    const result = await addAuditMasterData(catalyst, testData);
    
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

// Get all audit master data
app.get('/auditmaster', async (req, res) => {
  try {
    const catalyst = res.locals.catalyst;
    const { action, id, act, sector, recordCategory } = req.query;
    
    console.log(`Audit Master API Request - Action: ${action}, ID: ${id}, Act: ${act}, Sector: ${sector}, RecordCategory: ${recordCategory}`);
    
    let result;
    
    switch (action) {
      case 'getAll':
        result = await getAllAuditMasterData(catalyst);
        break;
        
      case 'getById':
        if (!id) {
          result = {
            status: 'error',
            message: 'ID parameter is required for getById action'
          };
        } else {
          result = await getAuditMasterById(catalyst, id);
        }
        break;
        
      case 'getByAct':
        if (!act) {
          result = {
            status: 'error',
            message: 'Act parameter is required for getByAct action'
          };
        } else {
          result = await getAuditMasterByAct(catalyst, act);
        }
        break;
        
      case 'getBySector':
        if (!sector) {
          result = {
            status: 'error',
            message: 'Sector parameter is required for getBySector action'
          };
        } else {
          result = await getAuditMasterBySector(catalyst, sector);
        }
        break;
        
      case 'getByRecordCategory':
        if (!recordCategory) {
          result = {
            status: 'error',
            message: 'RecordCategory parameter is required for getByRecordCategory action'
          };
        } else {
          result = await getAuditMasterByRecordCategory(catalyst, recordCategory);
        }
        break;
        
      case 'count':
        result = await getAuditMasterCount(catalyst);
        break;
        
      default:
        result = {
          status: 'error',
          message: 'Invalid action. Supported actions: getAll, getById, getByAct, getBySector, getByRecordCategory, count'
        };
    }
    
    res.status(result.status === 'success' ? 200 : 400).json(result);
    
  } catch (error) {
    console.error('Audit Master API Error:', error);
    res.status(500).json({
      status: 'error',
      message: 'Internal server error',
      error: error.message
    });
  }
});

// POST routes for data modification
app.post('/auditmaster', async (req, res) => {
  try {
    const catalyst = res.locals.catalyst;
    const { action } = req.query;
    const body = req.body;
    
    console.log(`Audit Master API POST Request - Action: ${action}`);
    
    let result;
    
    switch (action) {
      case 'add':
        if (!body) {
          result = {
            status: 'error',
            message: 'Request body is required for add action'
          };
        } else {
          result = await addAuditMasterData(catalyst, body);
        }
        break;
        
      case 'bulkImport':
        if (!body || !Array.isArray(body)) {
          result = {
            status: 'error',
            message: 'Request body must be an array for bulkImport action'
          };
        } else {
          result = await bulkImportAuditMasterData(catalyst, body);
        }
			break;
        
      case 'sendForApproval':
        if (!body || !body.itemIds || !Array.isArray(body.itemIds)) {
          result = {
            status: 'error',
            message: 'Request body must contain itemIds array for sendForApproval action'
          };
        } else {
          const siteName = body.siteName || '';
          result = await sendItemsForApproval(catalyst, body.itemIds, siteName);
        }
        break;
        
		default:
        result = {
          status: 'error',
          message: 'Invalid action. Supported actions: add, bulkImport, sendForApproval'
        };
    }
    
    res.status(result.status === 'success' ? 200 : 400).json(result);
    
  } catch (error) {
    console.error('Audit Master API Error:', error);
    res.status(500).json({
      status: 'error',
      message: 'Internal server error',
      error: error.message
    });
  }
});

// PUT route for updates
app.put('/auditmaster', async (req, res) => {
  try {
    const catalyst = res.locals.catalyst;
    const { action, id } = req.query;
    const body = req.body;
    
    console.log(`Audit Master API PUT Request - Action: ${action}, ID: ${id}`);
    
    let result;
    
    switch (action) {
      case 'update':
        if (!id || !body) {
          result = {
            status: 'error',
            message: 'ID and request body are required for update action'
          };
        } else {
          result = await updateAuditMasterData(catalyst, id, body);
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
    console.error('Audit Master API Error:', error);
    res.status(500).json({
      status: 'error',
      message: 'Internal server error',
      error: error.message
    });
  }
});

// DELETE route for deletions
app.delete('/auditmaster', async (req, res) => {
  try {
    const catalyst = res.locals.catalyst;
    const { action, id } = req.query;
    
    console.log(`Audit Master API DELETE Request - Action: ${action}, ID: ${id}`);
    
    let result;
    
    switch (action) {
      case 'delete':
        if (!id) {
          result = {
            status: 'error',
            message: 'ID parameter is required for delete action'
          };
        } else {
          result = await deleteAuditMasterData(catalyst, id);
        }
        break;
        
      case 'deleteAll':
        result = await deleteAllAuditMasterData(catalyst);
        break;
        
      default:
        result = {
          status: 'error',
          message: 'Invalid action. Supported actions: delete, deleteAll'
        };
    }
    
    res.status(result.status === 'success' ? 200 : 400).json(result);
    
  } catch (error) {
    console.error('Audit Master API Error:', error);
    res.status(500).json({
      status: 'error',
      message: 'Internal server error',
      error: error.message
    });
  }
});

module.exports = app;
