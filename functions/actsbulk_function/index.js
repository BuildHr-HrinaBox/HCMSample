'use strict';

const express = require('express');
const catalystSDK = require('zcatalyst-sdk-node');

const app = express();

app.use(express.json({ limit: '15mb' }));

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

// Table name for acts bulk data - make sure this matches exactly with your Data Store table name
// From the image, the table appears to be named "Actsbulk" (with capital A)
const TABLE_NAME = 'Actsbulk';

/** Data Store text columns expect strings; Excel often sends numbers, booleans, or rich-cell objects. */
const toStoreStr = (v) => {
  if (v === null || v === undefined) return '';
  if (typeof v === 'object') {
    if (v instanceof Date) return v.toISOString().slice(0, 10);
    if (v.v != null) return toStoreStr(v.v);
    if (v.w != null) return toStoreStr(v.w);
    return '';
  }
  if (typeof v === 'boolean') return v ? 'true' : 'false';
  return String(v).trim();
};

// Acts bulk data structure — column names must match Data Store exactly (see Actsbulk schema).
// This table has no Status column in Catalyst; do not send Status on insert (causes INVALID_INPUT).
const createActsBulkRecord = (data) => {
  return {
    Sector: toStoreStr(data.sector),
    Acts: toStoreStr(data.acts),
    Type: toStoreStr(data.type),
    States: toStoreStr(data.states),
    Description: toStoreStr(data.description),
    Applicability: toStoreStr(data.applicability),
    KeyComplianceRequirements: toStoreStr(data.keyComplianceRequirements),
    DueDate: toStoreStr(data.dueDate),
    PenaltyforNonCompliance: toStoreStr(data.penaltyforNonCompliance),
    Registers: toStoreStr(data.registers),
  };
};

// Convert Data Store record back to application format
const convertToAppFormat = (record) => {
  return {
    id: record.ROWID,
    sector: record.Sector || '',
    acts: record.Acts || '',
    type: record.Type || '',
    states: record.States || '',
    description: record.Description || '',
    applicability: record.Applicability || '',
    keyComplianceRequirements: record.KeyComplianceRequirements || '',
    dueDate: record.DueDate || '',
    penaltyforNonCompliance: record.PenaltyforNonCompliance || '',
    registers: record.Registers || '',
    status: record.Status || 'Yet to Start',
    createdTime: record.CREATEDTIME,
    modifiedTime: record.MODIFIEDTIME,
    creatorId: record.CREATORID
  };
};

// Get all acts bulk data
const getAllActsBulkData = async (catalyst) => {
  try {
    console.log('Fetching all acts bulk data from Data Store...');
    
    const dataStore = catalyst.datastore();
    const table = dataStore.table(TABLE_NAME);
    
    // Use getAllRows method instead of select
    const result = await table.getAllRows();
    
    console.log(`Found ${result.length} acts bulk records`);
    
    const formattedData = result.map(record => convertToAppFormat(record));
    
    return {
      status: 'success',
      message: 'Acts bulk data retrieved successfully',
      data: formattedData,
      count: formattedData.length
    };
  } catch (error) {
    console.error('Error fetching acts bulk data:', error);
    return {
      status: 'error',
      message: 'Failed to fetch acts bulk data',
      error: error.message
    };
  }
};

// Get acts bulk data by ID
const getActsBulkById = async (catalyst, id) => {
  try {
    console.log(`Fetching acts bulk data for ID: ${id}`);
    
    const dataStore = catalyst.datastore();
    const table = dataStore.table(TABLE_NAME);
    
    // Use getRow method instead of select
    const result = await table.getRow(id);
    
    if (!result) {
      return {
        status: 'error',
        message: 'Acts bulk record not found'
      };
    }
    
    const formattedData = convertToAppFormat(result);
    
    return {
      status: 'success',
      message: 'Acts bulk data retrieved successfully',
      data: formattedData
    };
  } catch (error) {
    console.error('Error fetching acts bulk data by ID:', error);
    return {
      status: 'error',
      message: 'Failed to fetch acts bulk data',
      error: error.message
    };
  }
};

// Add new acts bulk data
const addActsBulkData = async (catalyst, data) => {
  try {
    console.log('Adding new acts bulk data:', data);
    console.log('Table name being used:', TABLE_NAME);
    
    const dataStore = catalyst.datastore();
    const table = dataStore.table(TABLE_NAME);
    const record = createActsBulkRecord(data);
    
    console.log('Record to be inserted:', record);
    
    const result = await table.insertRow(record);
    
    console.log('Acts bulk data added successfully with ROWID:', result.ROWID);
    console.log('Full result object:', result);
    
    return {
      status: 'success',
      message: 'Acts bulk data added successfully',
      data: {
        id: result.ROWID,
        ...data
      }
    };
  } catch (error) {
    console.error('Error adding acts bulk data:', error);
    console.error('Error details:', error.message);
    console.error('Error stack:', error.stack);
    return {
      status: 'error',
      message: 'Failed to add acts bulk data',
      error: error.message
    };
  }
};

// Update acts bulk data
const updateActsBulkData = async (catalyst, id, data) => {
  try {
    console.log(`Updating acts bulk data for ID: ${id}`, data);
    
    const dataStore = catalyst.datastore();
    const table = dataStore.table(TABLE_NAME);
    const record = createActsBulkRecord(data);
    
    const result = await table.updateRow(id, record);
    
    console.log('Acts bulk data updated successfully');
    
    return {
      status: 'success',
      message: 'Acts bulk data updated successfully',
      data: {
        id: id,
        ...data
      }
    };
  } catch (error) {
    console.error('Error updating acts bulk data:', error);
    return {
      status: 'error',
      message: 'Failed to update acts bulk data',
      error: error.message
    };
  }
};

// Update only the status of an acts bulk record
const updateActsBulkStatus = async (catalyst, id, status) => {
  try {
    const dataStore = catalyst.datastore();
    const table = dataStore.table(TABLE_NAME);
    const row = await table.getRow(id);
    if (!row) {
      return { status: 'error', message: 'Acts bulk record not found' };
    }
    const updated = { ...row, Status: status };
    await table.updateRow(id, updated);
    return {
      status: 'success',
      message: 'Status updated successfully',
      data: { id, status }
    };
  } catch (error) {
    console.error('Error updating acts bulk status:', error);
    return {
      status: 'error',
      message: 'Failed to update status',
      error: error.message
    };
  }
};

// Delete acts bulk data
const deleteActsBulkData = async (catalyst, id) => {
  try {
    console.log(`Deleting acts bulk data for ID: ${id}`);
    
    const dataStore = catalyst.datastore();
    const table = dataStore.table(TABLE_NAME);
    await table.deleteRow(id);
    
    console.log('Acts bulk data deleted successfully');
    
    return {
      status: 'success',
      message: 'Acts bulk data deleted successfully'
    };
  } catch (error) {
    console.error('Error deleting acts bulk data:', error);
    return {
      status: 'error',
      message: 'Failed to delete acts bulk data',
      error: error.message
    };
  }
};

// Delete all acts bulk data (bulk delete)
const deleteAllActsBulkData = async (catalyst) => {
  try {
    console.log('Deleting all acts bulk data from Data Store...');
    
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
        message: `Successfully deleted ${deletedCount} acts bulk records`,
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
    console.error('Error deleting all acts bulk data:', error);
    console.error('Error stack:', error.stack);
    return {
      status: 'error',
      message: 'Failed to delete all acts bulk data',
      error: error.message || error.toString()
    };
  }
};

const BULK_INSERT_BATCH_SIZE = 100;

function formatBulkImportError(error) {
  const code = error && (error.code || error.errorInfo?.code);
  const msg = (error && error.message) || String(error);
  const parts = [msg];
  if (code) parts.push(`code: ${code}`);
  return parts.join(' — ');
}

// Bulk import acts data
const bulkImportActsData = async (catalyst, dataArray) => {
  try {
    console.log(`Bulk importing ${dataArray.length} acts bulk records`);
    console.log('Table name being used:', TABLE_NAME);

    const dataStore = catalyst.datastore();
    const table = dataStore.table(TABLE_NAME);
    const records = dataArray.map((data) => createActsBulkRecord(data));

    console.log('Records to be inserted (sample keys):', records.slice(0, 1));

    const allInserted = [];
    for (let offset = 0; offset < records.length; offset += BULK_INSERT_BATCH_SIZE) {
      const chunk = records.slice(offset, offset + BULK_INSERT_BATCH_SIZE);
      const chunkSource = dataArray.slice(offset, offset + BULK_INSERT_BATCH_SIZE);
      console.log(`Inserting batch ${Math.floor(offset / BULK_INSERT_BATCH_SIZE) + 1}, rows ${offset + 1}-${offset + chunk.length}`);
      const batchResult = await table.insertRows(chunk);
      if (!Array.isArray(batchResult)) {
        console.error('insertRows returned non-array:', batchResult);
        return {
          status: 'error',
          message: 'Bulk import returned an unexpected response from Data Store',
          insertedCount: allInserted.length,
          error: 'insertRows did not return an array',
        };
      }
      batchResult.forEach((row, i) => {
        allInserted.push({
          id: row.ROWID,
          ...chunkSource[i],
        });
      });
    }

    console.log(`Bulk import completed. ${allInserted.length} records added`);

    return {
      status: 'success',
      message: `Bulk import completed successfully. ${allInserted.length} records added`,
      insertedCount: allInserted.length,
      data: allInserted,
    };
  } catch (error) {
    console.error('Error in bulk import:', error);
    console.error('Error details:', error && error.message, error && error.code);
    console.error('Error stack:', error && error.stack);
    const detail = formatBulkImportError(error);
    return {
      status: 'error',
      message: `Failed to bulk import acts data: ${detail}`,
      error: detail,
    };
  }
};

// Get acts data by sector
const getActsDataBySector = async (catalyst, sector) => {
  try {
    console.log(`Fetching acts data for sector: ${sector}`);
    
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
      message: `Acts data for sector '${sector}' retrieved successfully`,
      data: formattedData,
      count: formattedData.length
    };
  } catch (error) {
    console.error('Error fetching acts data by sector:', error);
    return {
      status: 'error',
      message: 'Failed to fetch acts data by sector',
      error: error.message
    };
  }
};

// Get acts data by type
const getActsDataByType = async (catalyst, type) => {
  try {
    console.log(`Fetching acts data for type: ${type}`);
    
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
      message: `Acts data for type '${type}' retrieved successfully`,
      data: formattedData,
      count: formattedData.length
    };
  } catch (error) {
    console.error('Error fetching acts data by type:', error);
    return {
      status: 'error',
      message: 'Failed to fetch acts data by type',
      error: error.message
    };
  }
};

// Get acts data by states
const getActsDataByStates = async (catalyst, states) => {
  try {
    console.log(`Fetching acts data for states: ${states}`);
    
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
      message: `Acts data for states '${states}' retrieved successfully`,
      data: formattedData,
      count: formattedData.length
    };
  } catch (error) {
    console.error('Error fetching acts data by states:', error);
    return {
      status: 'error',
      message: 'Failed to fetch acts data by states',
      error: error.message
    };
  }
};

// Get count of acts bulk records
const getActsBulkCount = async (catalyst) => {
  try {
    console.log('Getting acts bulk data count...');
    
    const dataStore = catalyst.datastore();
    const table = dataStore.table(TABLE_NAME);
    
    // Use getAllRows and get length
    const result = await table.getAllRows();
    
    return {
      status: 'success',
      message: 'Acts bulk data count retrieved successfully',
      count: result.length
    };
  } catch (error) {
    console.error('Error getting acts bulk count:', error);
    return {
      status: 'error',
      message: 'Failed to get acts bulk data count',
      error: error.message
    };
  }
};

// Populate table with sample data
const populateTableWithSampleData = async (catalyst) => {
  try {
    console.log('Populating Actsbulk table with sample data...');
    
    const sampleData = [
      {
        Sector: "Manufacturing",
        Acts: "Factories Act, 1948",
        Type: "Central Act",
        States: "All States",
        Description: "An Act to consolidate and amend the law regulating labour in factories.",
        Applicability: "All factories employing 10 or more workers with power, or 20 or more workers without power",
        KeyComplianceRequirements: "Registration, Safety measures, Working hours, Health facilities",
        DueDate: "Within 15 days of establishment",
        PenaltyforNonCompliance: "Fine up to Rs. 2 lakhs or imprisonment up to 2 years",
        Registers: "Muster roll, Register of workers, Overtime register, Leave register"
      },
      {
        Sector: "Manufacturing",
        Acts: "Industrial Disputes Act, 1947",
        Type: "Central Act",
        States: "All States",
        Description: "An Act to make provision for the investigation and settlement of industrial disputes.",
        Applicability: "All industrial establishments",
        KeyComplianceRequirements: "Notice of change, Standing orders, Lay-off compensation",
        DueDate: "As per notice period specified",
        PenaltyforNonCompliance: "Fine up to Rs. 1,000 or imprisonment up to 3 months",
        Registers: "Standing orders register, Notice register"
      },
      {
        Sector: "Healthcare",
        Acts: "Clinical Establishments Act, 2010",
        Type: "Central Act",
        States: "All States",
        Description: "An Act to provide for registration and regulation of clinical establishments.",
        Applicability: "All clinical establishments",
        KeyComplianceRequirements: "Registration, Minimum standards, Display of rates",
        DueDate: "Within 30 days of establishment",
        PenaltyforNonCompliance: "Fine up to Rs. 5 lakhs",
        Registers: "Patient register, Medical records register"
      },
      {
        Sector: "Finance",
        Acts: "Banking Regulation Act, 1949",
        Type: "Central Act",
        States: "All States",
        Description: "An Act to consolidate and amend the law relating to banking.",
        Applicability: "All banking companies",
        KeyComplianceRequirements: "Licensing, Capital requirements, Reserve requirements",
        DueDate: "As per RBI guidelines",
        PenaltyforNonCompliance: "Fine up to Rs. 1 crore or imprisonment",
        Registers: "Account books, Transaction register"
      },
      {
        Sector: "Environment",
        Acts: "Environment Protection Act, 1986",
        Type: "Central Act",
        States: "All States",
        Description: "An Act to provide for the protection and improvement of environment.",
        Applicability: "All industries and establishments",
        KeyComplianceRequirements: "Environmental clearance, Pollution control measures",
        DueDate: "Before commencement of operations",
        PenaltyforNonCompliance: "Fine up to Rs. 1 lakh or imprisonment up to 5 years",
        Registers: "Environmental monitoring register"
      },
      {
        Sector: "Manufacturing",
        Acts: "Maharashtra Factories Rules, 1963",
        Type: "State Rules",
        States: "Maharashtra",
        Description: "Rules framed under the Factories Act, 1948 for Maharashtra state.",
        Applicability: "All factories in Maharashtra",
        KeyComplianceRequirements: "State-specific safety measures, Local compliance",
        DueDate: "As per state notification",
        PenaltyforNonCompliance: "As per state rules",
        Registers: "State-specific registers as required"
      },
      {
        Sector: "Manufacturing",
        Acts: "Karnataka Factories Rules, 1969",
        Type: "State Rules",
        States: "Karnataka",
        Description: "Rules framed under the Factories Act, 1948 for Karnataka state.",
        Applicability: "All factories in Karnataka",
        KeyComplianceRequirements: "State-specific safety measures, Local compliance",
        DueDate: "As per state notification",
        PenaltyforNonCompliance: "As per state rules",
        Registers: "State-specific registers as required"
      },
      {
        Sector: "Healthcare",
        Acts: "Tamil Nadu Clinical Establishments Rules, 2018",
        Type: "State Rules",
        States: "Tamil Nadu",
        Description: "Rules framed under the Clinical Establishments Act, 2010 for Tamil Nadu state.",
        Applicability: "All clinical establishments in Tamil Nadu",
        KeyComplianceRequirements: "State registration, Local standards",
        DueDate: "Within 30 days of establishment",
        PenaltyforNonCompliance: "As per state rules",
        Registers: "State-specific medical records register"
      }
    ];
    
    const dataStore = catalyst.datastore();
    const table = dataStore.table(TABLE_NAME);
    const result = await table.insertRows(sampleData);
    
    console.log(`Successfully populated table with ${result.length} sample records`);
    
    return {
      status: 'success',
      message: `Successfully populated Actsbulk table with ${result.length} sample records`,
      data: result.map((record, index) => ({
        id: record.ROWID,
        sector: sampleData[index].Sector,
        acts: sampleData[index].Acts,
        type: sampleData[index].Type,
        states: sampleData[index].States,
        description: sampleData[index].Description,
        applicability: sampleData[index].Applicability,
        keyComplianceRequirements: sampleData[index].KeyComplianceRequirements,
        dueDate: sampleData[index].DueDate,
        penaltyforNonCompliance: sampleData[index].PenaltyforNonCompliance,
        registers: sampleData[index].Registers
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
  res.status(200).json({ status: 'success', message: 'actsbulk_function ready' });
});

// Test endpoint to check table access and verify data
app.get('/test-table', async (req, res) => {
  try {
    console.log('Testing Actsbulk table access...');
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
      acts: "Test Act",
      type: "Test Type",
      states: "Test State",
      description: "Test Description",
      applicability: "Test Applicability",
      keyComplianceRequirements: "Test Requirements",
      dueDate: "2024-12-31",
      penaltyforNonCompliance: "Test Penalty",
      registers: "Test Registers"
    };
    
    console.log('Test data to be added:', testData);
    
    const result = await addActsBulkData(catalyst, testData);
    
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

// Get all acts bulk data
app.get('/actsbulk', async (req, res) => {
  try {
    const catalyst = res.locals.catalyst;
    const { action, id, sector, type, states } = req.query;
    
    console.log(`Acts Bulk API Request - Action: ${action}, ID: ${id}, Sector: ${sector}, Type: ${type}, States: ${states}`);
    
    let result;
    
    switch (action) {
      case 'getAll':
        result = await getAllActsBulkData(catalyst);
        break;
        
      case 'getById':
        if (!id) {
          result = {
            status: 'error',
            message: 'ID parameter is required for getById action'
          };
        } else {
          result = await getActsBulkById(catalyst, id);
        }
        break;
        
      case 'getBySector':
        if (!sector) {
          result = {
            status: 'error',
            message: 'Sector parameter is required for getBySector action'
          };
        } else {
          result = await getActsDataBySector(catalyst, sector);
        }
        break;
        
      case 'getByType':
        if (!type) {
          result = {
            status: 'error',
            message: 'Type parameter is required for getByType action'
          };
        } else {
          result = await getActsDataByType(catalyst, type);
        }
        break;
        
      case 'getByStates':
        if (!states) {
          result = {
            status: 'error',
            message: 'States parameter is required for getByStates action'
          };
        } else {
          result = await getActsDataByStates(catalyst, states);
        }
        break;
        
      case 'count':
        result = await getActsBulkCount(catalyst);
        break;
        
      case 'populate':
        result = await populateTableWithSampleData(catalyst);
        break;
        
      default:
        result = {
          status: 'error',
          message: 'Invalid action. Supported actions: getAll, getById, getBySector, getByType, getByStates, count, populate'
        };
    }
    
    res.status(result.status === 'success' ? 200 : 400).json(result);
    
  } catch (error) {
    console.error('Acts Bulk API Error:', error);
    res.status(500).json({
      status: 'error',
      message: 'Internal server error',
      error: error.message
    });
  }
});

// POST routes for data modification
app.post('/actsbulk', async (req, res) => {
  try {
    const catalyst = res.locals.catalyst;
    const { action } = req.query;
    const body = req.body;
    
    console.log(`Acts Bulk API POST Request - Action: ${action}`);
    
    let result;
    
    switch (action) {
      case 'add':
        if (!body) {
          result = {
            status: 'error',
            message: 'Request body is required for add action'
          };
        } else {
          result = await addActsBulkData(catalyst, body);
        }
        break;
        
      case 'bulkImport':
        if (!body || !Array.isArray(body)) {
          result = {
            status: 'error',
            message: 'Request body must be an array for bulkImport action'
          };
        } else {
          result = await bulkImportActsData(catalyst, body);
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
    console.error('Acts Bulk API Error:', error);
    res.status(500).json({
      status: 'error',
      message: 'Internal server error',
      error: error.message
    });
  }
});

// PUT route for updates
app.put('/actsbulk', async (req, res) => {
  try {
    const catalyst = res.locals.catalyst;
    const { action, id } = req.query;
    const body = req.body;
    
    console.log(`Acts Bulk API PUT Request - Action: ${action}, ID: ${id}`);
    
    let result;
    
    switch (action) {
      case 'update':
        if (!id || !body) {
          result = {
            status: 'error',
            message: 'ID and request body are required for update action'
          };
        } else {
          result = await updateActsBulkData(catalyst, id, body);
        }
        break;
      case 'updateStatus':
        if (!id || body == null || body.status == null) {
          result = {
            status: 'error',
            message: 'ID and body.status are required for updateStatus action'
          };
        } else {
          result = await updateActsBulkStatus(catalyst, id, body.status);
        }
        break;
      default:
        result = {
          status: 'error',
          message: 'Invalid action. Supported actions: update, updateStatus'
        };
    }
    
    res.status(result.status === 'success' ? 200 : 400).json(result);
    
  } catch (error) {
    console.error('Acts Bulk API Error:', error);
    res.status(500).json({
      status: 'error',
      message: 'Internal server error',
      error: error.message
    });
  }
});

// DELETE route for deletions
app.delete('/actsbulk', async (req, res) => {
  try {
    const catalyst = res.locals.catalyst;
    const { action, id } = req.query;
    
    console.log(`Acts Bulk API DELETE Request - Action: ${action}, ID: ${id}`);
    
    let result;
    
    switch (action) {
      case 'delete':
        if (!id) {
          result = {
            status: 'error',
            message: 'ID parameter is required for delete action'
          };
        } else {
          result = await deleteActsBulkData(catalyst, id);
        }
        break;
        
      case 'deleteAll':
        result = await deleteAllActsBulkData(catalyst);
        break;
        
      default:
        result = {
          status: 'error',
          message: 'Invalid action. Supported actions: delete, deleteAll'
        };
    }
    
    res.status(result.status === 'success' ? 200 : 400).json(result);
    
  } catch (error) {
    console.error('Acts Bulk API Error:', error);
    res.status(500).json({
      status: 'error',
      message: 'Internal server error',
      error: error.message
    });
  }
});

module.exports = app;
