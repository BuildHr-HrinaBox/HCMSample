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

// Table name for checklist bulk data
const TABLE_NAME = 'checklistbulk';
const FORM_MASTER_TABLE = 'FormMaster';

function squashKeyPart(v) {
  return String(v ?? '')
    .trim()
    .toLowerCase()
    .replace(/\.(xlsx|xls)$/i, '')
    .replace(/\s+/g, ' ');
}

/** Same key as Form Master / Statutory — one template per form+act+description+sector+state (+ Form U variant). */
function buildChecklistBulkMatchKey(data) {
  const base = [
    squashKeyPart(data.formName || data.FormName),
    squashKeyPart(data.act || data.Act),
    squashKeyPart(data.description || data.Description),
    squashKeyPart(data.sector || data.Sector) || 'nosector',
    squashKeyPart(data.state || data.State) || 'nostate'
  ].join('|');
  const variant = inferBulkFormVariantFromRow(data);
  return variant ? `${base}|${variant}` : base;
}

function buildChecklistBulkIdentityKey(data) {
  const norm = (v) =>
    String(v ?? '')
      .trim()
      .toLowerCase()
      .replace(/\s+/g, ' ');
  return [
    norm(data.formName || data.FormName),
    norm(data.sector || data.Sector),
    norm(data.state || data.State),
    norm(data.act || data.Act),
    norm(data.description || data.Description)
  ].join('\x1f');
}

function buildDisplayTemplateFileName(record) {
  const original = String(record?.TemplateFileName || record?.templateFileName || '').trim();
  const formName = String(record?.FormName || record?.formName || 'Form').trim();
  const state = String(record?.State || record?.state || '').trim();
  const act = String(record?.Act || record?.act || '').trim();
  const looksGeneric =
    !original ||
    /^form\s*[a-z0-9]*\.xlsx$/i.test(original) ||
    original.toLowerCase() === `${formName.toLowerCase()}.xlsx`;
  if (!looksGeneric) return original;
  const parts = [formName];
  if (state) parts.push(state);
  else if (act) parts.push(act.slice(0, 40));
  const ext = original.includes('.') ? original.slice(original.lastIndexOf('.')) : '.xlsx';
  return `${parts.join(' - ')}${ext}`.replace(/[/\\?%*:|"<>]/g, '_');
}

function inferFormUVariantFromBulkRow(data) {
  const blob = [
    data.description,
    data.Description,
    data.act,
    data.Act,
    data.formName,
    data.FormName,
    data.formFileName,
    data.FormFileName
  ]
    .filter(Boolean)
    .join(' ')
    .toLowerCase();
  if (/confidential|rule\s*34|confidential\s+character/.test(blob)) {
    return 'confidential_character';
  }
  if (
    /employee|particulars|register\s+of\s+employees|shops\s+and\s+establishment|wage\s+register/.test(
      blob
    )
  ) {
    return 'employee_register';
  }
  return null;
}

function inferFormXVariantFromBulkRow(data) {
  const blob = [
    data.description,
    data.Description,
    data.act,
    data.Act,
    data.state,
    data.State,
    data.sector,
    data.Sector,
    data.formName,
    data.FormName,
    data.formFileName,
    data.FormFileName
  ]
    .filter(Boolean)
    .join(' ')
    .toLowerCase();
  if (/identity\s+card/i.test(blob)) return 'identity_card';
  if (/register\s+of\s+fines|fines\s+register|fine\s+register/i.test(blob)) {
    return 'register_of_fines';
  }
  if (/leave\s+with\s+wages|leave\s+register|register\s+of\s+leave|earned\s+leave/i.test(blob)) {
    return 'leave_register';
  }
  if (
    /\bform\s*[-"']?\s*x\b/i.test(blob) &&
    /andhra\s+pradesh|telangana|\bap\b/.test(blob) &&
    /shops?\s*(and|&)\s*establishment/i.test(blob) &&
    !/identity\s+card|leave\s+with\s+wages|leave\s+register|register\s+of\s+leave|earned\s+leave/i.test(blob)
  ) {
    return 'register_of_fines';
  }
  return null;
}

function inferBulkFormVariantFromRow(data) {
  return inferFormUVariantFromBulkRow(data) || inferFormXVariantFromBulkRow(data) || null;
}

/** When checklist row has State, template must use the same State (not another state's file). */
function statesCompatibleForFormTemplateLink(bulkState, masterState) {
  const b = squashKeyPart(bulkState);
  const m = squashKeyPart(masterState);
  if (b && m) return b === m;
  if (b && !m) return false;
  return true;
}

function scoreBulkToFormMasterLine(bulkLine, masterLine) {
  let score = 0;
  if (squashKeyPart(bulkLine.formName) !== squashKeyPart(masterLine.formName)) return -999;
  if (!statesCompatibleForFormTemplateLink(bulkLine.state, masterLine.state)) return -999;
  if (squashKeyPart(bulkLine.act) && squashKeyPart(bulkLine.act) === squashKeyPart(masterLine.act)) {
    score += 5;
  }
  if (squashKeyPart(bulkLine.state) && squashKeyPart(bulkLine.state) === squashKeyPart(masterLine.state)) {
    score += 5;
  }
  if (squashKeyPart(bulkLine.sector) && squashKeyPart(bulkLine.sector) === squashKeyPart(masterLine.sector)) {
    score += 3;
  }
  const bd = squashKeyPart(bulkLine.description);
  const md = squashKeyPart(masterLine.description);
  if (bd && md) {
    if (bd === md) score += 12;
    else if (bd.includes(md) || md.includes(bd)) score += 7;
    else score -= 40;
  }
  const bv = inferBulkFormVariantFromRow(bulkLine);
  const mv = inferBulkFormVariantFromRow(masterLine);
  if (bv && mv) {
    if (bv === mv) score += 25;
    else score -= 200;
  }
  return score;
}

async function buildFormMasterFileLookup(catalyst) {
  const lookup = new Map();
  const allTemplates = [];
  try {
    const table = catalyst.datastore().table(FORM_MASTER_TABLE);
    const rows = await table.getAllRows();
    (rows || []).forEach((record) => {
      const fileId = record.Action != null ? String(record.Action).trim() : '';
      if (!fileId) return;
      const line = {
        formName: record.FormName || '',
        act: record.Act || '',
        description: record.Description || '',
        sector: record.Sector || '',
        state: record.State || ''
      };
      allTemplates.push({
        line,
        payload: {
          formFile: fileId,
          formFileName: buildDisplayTemplateFileName({
            FormName: line.formName,
            State: line.state,
            Act: line.act,
            TemplateFileName: record.TemplateFileName || ''
          })
        }
      });
    });
    allTemplates.forEach(({ line, payload }) => {
      const keys = new Set([
        buildChecklistBulkMatchKey(line),
        buildChecklistBulkIdentityKey(line)
      ].filter(Boolean));
      keys.forEach((key) => lookup.set(key, payload));
    });
  } catch (err) {
    console.warn('Form Master lookup skipped for checklistbulk:', err.message);
  }
  lookup.__allTemplates = allTemplates;
  return lookup;
}

function resolveFormFileForBulkRecordScored(data, lookup) {
  if (!lookup) {
    return { formFile: '', formFileName: '' };
  }
  const line = {
    formName: data.formName || data.FormName || '',
    act: data.act || data.Act || '',
    description: data.description || data.Description || '',
    sector: data.sector || data.Sector || '',
    state: data.state || data.State || ''
  };
  const exact =
    lookup.get(buildChecklistBulkMatchKey(line)) ||
    lookup.get(buildChecklistBulkIdentityKey(line));
  if (exact?.formFile) return exact;
  const templates = lookup.__allTemplates || [];
  if (!templates.length) {
    return { formFile: '', formFileName: '' };
  }
  let best = null;
  let bestScore = -Infinity;
  templates.forEach(({ line: masterLine, payload }) => {
    const s = scoreBulkToFormMasterLine(line, masterLine);
    if (s > bestScore) {
      bestScore = s;
      best = payload;
    }
  });
  if (best && bestScore > 0) {
    return best;
  }
  return { formFile: '', formFileName: '' };
}

function resolveFormFileForBulkRecord(data, lookup) {
  return resolveFormFileForBulkRecordScored(data, lookup);
}

const isMonthlyFrequency = (frequency) =>
  String(frequency || '').toLowerCase().includes('monthly');

/** Day-of-month (1–31) from due date; recovers mistaken Excel serial → 1900-01-* rows. */
const extractMonthlyDayOfMonth = (value) => {
  if (value == null || value === '') return null;
  if (typeof value === 'number' && Number.isInteger(value) && value >= 1 && value <= 31) {
    return value;
  }
  const s = String(value).trim();
  if (!s) return null;
  if (/^\d{1,2}$/.test(s)) {
    const day = parseInt(s, 10);
    if (day >= 1 && day <= 31) return day;
  }
  const legacy = s.match(/^1900-01-(\d{1,2})$/);
  if (legacy) {
    const recovered = parseInt(legacy[1], 10) + 1;
    if (recovered >= 1 && recovered <= 31) return recovered;
  }
  return null;
};

// Normalize incoming DueDate: day-of-month (1–31) for monthly, ISO for real dates, else text
const normalizeToISODate = (value, frequency = '') => {
  if (!value && value !== 0) return '';
  const monthly = isMonthlyFrequency(frequency);
  const pad2 = (n) => String(n).padStart(2, '0');
  const fmt = (d) => `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())}`;

  if (typeof value === 'number') {
    if (Number.isInteger(value) && value >= 1 && value <= 31) {
      return String(value);
    }
    if (value > 31) {
      const excelEpoch = new Date(1900, 0, 1);
      const jsDate = new Date(excelEpoch.getTime() + (value - 2) * 24 * 60 * 60 * 1000);
      return fmt(jsDate);
    }
    return String(value);
  }

  if (typeof value === 'string') {
    const s = value.trim();
    if (!s) return '';

    const monthlyDay = extractMonthlyDayOfMonth(s);
    if (monthlyDay != null && monthly) {
      return String(monthlyDay);
    }
    if (/^\d{1,2}$/.test(s)) {
      const day = parseInt(s, 10);
      if (day >= 1 && day <= 31) return String(day);
    }

    // Try to parse as ISO date format (YYYY-MM-DD or YYYY/MM/DD)
    const iso = s.match(/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})$/);
    if (iso) {
      const y = Number(iso[1]);
      const m = Number(iso[2]);
      const d = Number(iso[3]);
      if (y === 1900 && monthly) {
        const recovered = extractMonthlyDayOfMonth(s);
        if (recovered != null) return String(recovered);
      }
      // Validate the date
      const dateObj = new Date(y, m - 1, d);
      if (dateObj.getFullYear() === y && dateObj.getMonth() === m - 1 && dateObj.getDate() === d) {
        return `${y}-${pad2(m)}-${pad2(d)}`;
      }
    }
    
    // Try to parse as date with month name (DD-MMM-YYYY or DD MMM YYYY)
    const mon = { jan:0,feb:1,mar:2,apr:3,may:4,jun:5,jul:6,aug:7,sep:8,sept:8,oct:9,nov:10,dec:11 };
    const dmy = s.match(/^(\d{1,2})[-\s\/]+([A-Za-z]{3,4})[-\s\/]+(\d{4})$/);
    if (dmy) {
      const d = Number(dmy[1]);
      const mIdx = mon[dmy[2].toLowerCase()];
      const y = Number(dmy[3]);
      if (mIdx !== undefined) {
        const dateObj = new Date(y, mIdx, d);
        if (dateObj.getFullYear() === y && dateObj.getMonth() === mIdx && dateObj.getDate() === d) {
          return fmt(dateObj);
        }
      }
    }
    
    // Try to parse as a general date
    const parsed = new Date(s);
    if (!isNaN(parsed.getTime())) {
      // Check if it's a valid date (not just a number that was parsed)
      if (s.match(/^\d+$/) && s.length > 4) {
        // It's likely an Excel serial number as string
        return fmt(parsed);
      } else if (parsed.getFullYear() > 1900 && parsed.getFullYear() < 2100) {
        // It's a reasonable date - verify it's actually a date and not just parsed text
        // Check if the original string looks like a date format
        if (s.match(/\d{1,2}[-\/\s]\d{1,2}[-\/\s]\d{2,4}/) || s.match(/\d{4}[-\/]\d{1,2}[-\/]\d{1,2}/)) {
          return fmt(parsed);
        }
      }
    }
    
    // If we can't parse it as a date, return the original string value
    // This handles cases like "Monthly Basis", "Quarterly", etc.
    // Since DueDate is now a text column, we can store these values directly
    return s;
  }
  if (value instanceof Date) return fmt(value);
  return '';
};

// Checklist bulk data structure based on the Data Store schema
const createChecklistBulkRecord = (data, formFileFields = null) => {
  const normalizedDueDate = normalizeToISODate(data.dueDate, data.frequency);
  const resolved =
    formFileFields ||
    resolveFormFileForBulkRecord(data, data.__formMasterLookup || null);

  return {
    Sector: data.sector || '',
    State: data.state || '',
    Act: data.act || '',
    FormName: data.formName || '',
    ConcernedGovtDepartment: data.concernedGovtDepartment || '',
    DueDate: normalizedDueDate,
    Description: data.description || '',
    Nameofthecode: data.nameOfTheCode || data.nameofthecode || '',
    Frequency: data.frequency || '',
    NameoftheRule: data.nameOfTheRule || data.nameoftheRule || '',
    FormFile: resolved.formFile || data.formFile || data.FormFile || ''
  };
};

// Convert Data Store record back to application format
const convertToAppFormat = (record) => {
  return {
    id: record.ROWID,
    sector: record.Sector || '',
    state: record.State || '',
    act: record.Act || '',
    formName: record.FormName || '',
    concernedGovtDepartment: record.ConcernedGovtDepartment || '',
    dueDate: record.DueDate || '',
    description: record.Description || '',
    nameOfTheCode: record.Nameofthecode || '',
    frequency: record.Frequency || '',
    nameOfTheRule: record.NameoftheRule || '',
    formFile: record.FormFile || null,
    formFileName:
      record.FormFileName ||
      (record.FormFile
        ? buildDisplayTemplateFileName({
            FormName: record.FormName,
            State: record.State,
            Act: record.Act,
            TemplateFileName: ''
          })
        : null),
    matchKey: buildChecklistBulkMatchKey({
      formName: record.FormName,
      act: record.Act,
      description: record.Description,
      sector: record.Sector,
      state: record.State
    }),
    createdTime: record.CREATEDTIME,
    modifiedTime: record.MODIFIEDTIME,
    creatorId: record.CREATORID
  };
};

// Get all checklist bulk data
const getAllChecklistBulkData = async (catalyst) => {
  try {
    console.log('Fetching all checklist bulk data from Data Store...');
    
    const dataStore = catalyst.datastore();
    const table = dataStore.table(TABLE_NAME);
    
    // Use getAllRows method instead of select
    const result = await table.getAllRows();
    
    console.log(`Found ${result.length} checklist bulk records`);
    
    const formattedData = result.map(record => convertToAppFormat(record));
    
    return {
      status: 'success',
      message: 'Checklist bulk data retrieved successfully',
      data: formattedData,
      count: formattedData.length
    };
  } catch (error) {
    console.error('Error fetching checklist bulk data:', error);
    return {
      status: 'error',
      message: 'Failed to fetch checklist bulk data',
      error: error.message
    };
  }
};

// Get checklist bulk data by ID
const getChecklistBulkById = async (catalyst, id) => {
  try {
    console.log(`Fetching checklist bulk data for ID: ${id}`);
    
    const dataStore = catalyst.datastore();
    const table = dataStore.table(TABLE_NAME);
    
    // Use getRow method instead of select
    const result = await table.getRow(id);
    
    if (!result) {
      return {
        status: 'error',
        message: 'Checklist bulk record not found'
      };
    }
    
    const formattedData = convertToAppFormat(result);
    
    return {
      status: 'success',
      message: 'Checklist bulk data retrieved successfully',
      data: formattedData
    };
  } catch (error) {
    console.error('Error fetching checklist bulk data by ID:', error);
    return {
      status: 'error',
      message: 'Failed to fetch checklist bulk data',
      error: error.message
    };
  }
};

// Add new checklist bulk data
const addChecklistBulkData = async (catalyst, data) => {
  try {
    console.log('Adding new checklist bulk data:', data);
    console.log('Table name being used:', TABLE_NAME);
    
    const dataStore = catalyst.datastore();
    const table = dataStore.table(TABLE_NAME);
    const formMasterLookup = await buildFormMasterFileLookup(catalyst);
    const record = createChecklistBulkRecord(data, resolveFormFileForBulkRecord(data, formMasterLookup));
    
    console.log('Record to be inserted:', record);
    
    const result = await table.insertRow(record);
    
    console.log('Checklist bulk data added successfully with ROWID:', result.ROWID);
    console.log('Full result object:', result);
    
    return {
      status: 'success',
      message: 'Checklist bulk data added successfully',
      data: {
        id: result.ROWID,
        ...data
      }
    };
  } catch (error) {
    console.error('Error adding checklist bulk data:', error);
    console.error('Error details:', error.message);
    console.error('Error stack:', error.stack);
    return {
      status: 'error',
      message: 'Failed to add checklist bulk data',
      error: error.message
    };
  }
};

// Update checklist bulk data
const updateChecklistBulkData = async (catalyst, id, data) => {
  try {
    console.log(`Updating checklist bulk data for ID: ${id}`, data);
    
    const dataStore = catalyst.datastore();
    const table = dataStore.table(TABLE_NAME);
    const formMasterLookup = await buildFormMasterFileLookup(catalyst);
    const record = createChecklistBulkRecord(data, resolveFormFileForBulkRecord(data, formMasterLookup));
    
    const result = await table.updateRow(id, record);
    
    console.log('Checklist bulk data updated successfully');
    
    return {
      status: 'success',
      message: 'Checklist bulk data updated successfully',
      data: {
        id: id,
        ...data
      }
    };
  } catch (error) {
    console.error('Error updating checklist bulk data:', error);
    return {
      status: 'error',
      message: 'Failed to update checklist bulk data',
      error: error.message
    };
  }
};

// Delete checklist bulk data
const deleteChecklistBulkData = async (catalyst, id) => {
  try {
    console.log(`Deleting checklist bulk data for ID: ${id}`);
    
    const dataStore = catalyst.datastore();
    const table = dataStore.table(TABLE_NAME);
    await table.deleteRow(id);
    
    console.log('Checklist bulk data deleted successfully');
    
    return {
      status: 'success',
      message: 'Checklist bulk data deleted successfully'
    };
  } catch (error) {
    console.error('Error deleting checklist bulk data:', error);
    return {
      status: 'error',
      message: 'Failed to delete checklist bulk data',
      error: error.message
    };
  }
};

// Helper function to delete Checklist entries by formName
const deleteChecklistEntriesByFormNames = async (catalyst, formNames) => {
  if (!formNames || formNames.length === 0) {
    return { deletedCount: 0, errors: [] };
  }

  try {
    console.log(`Deleting Checklist entries for ${formNames.length} formName(s)...`);
    const zcql = catalyst.zcql();
    const checklistTable = catalyst.datastore().table('Checklist');
    
    let deletedCount = 0;
    const errors = [];
    
    // Get all Checklist entries and filter by formName
    const checklistRows = await zcql.executeZCQLQuery(
      'SELECT ROWID, FormName FROM Checklist'
    );
    
    // Create a set of formNames for quick lookup (case-insensitive)
    const formNameSet = new Set(formNames.map(name => String(name).toLowerCase().trim()));
    
    // Delete matching Checklist entries
    for (const row of checklistRows) {
      const checklistFormName = row.Checklist?.FormName;
      if (checklistFormName && formNameSet.has(String(checklistFormName).toLowerCase().trim())) {
        try {
          const rowId = row.Checklist.ROWID;
          await checklistTable.deleteRow(rowId);
          deletedCount++;
          console.log(`Deleted Checklist entry with formName: ${checklistFormName}, ROWID: ${rowId}`);
        } catch (deleteError) {
          console.error(`Error deleting Checklist entry ${row.Checklist.ROWID}:`, deleteError);
          errors.push({
            formName: checklistFormName,
            rowId: row.Checklist.ROWID,
            error: deleteError.message
          });
        }
      }
    }
    
    console.log(`Deleted ${deletedCount} Checklist entries, ${errors.length} errors`);
    return { deletedCount, errors };
  } catch (error) {
    console.error('Error deleting Checklist entries:', error);
    return { deletedCount: 0, errors: [{ error: error.message }] };
  }
};

// Bulk delete all checklist bulk data
const bulkDeleteChecklistData = async (catalyst) => {
  try {
    console.log('Bulk deleting all checklist bulk data...');
    
    const dataStore = catalyst.datastore();
    const table = dataStore.table(TABLE_NAME);
    const zcql = catalyst.zcql();
    
    // First get all rows with formNames before deleting
    let allRows;
    let formNames = [];
    try {
      allRows = await table.getAllRows();
      console.log(`Found ${allRows.length} records to delete`);
      
      // Extract formNames from all rows
      formNames = allRows
        .map(row => row.FormName || row[TABLE_NAME]?.FormName)
        .filter(name => name && String(name).trim());
      
      console.log(`Extracted ${formNames.length} unique formNames:`, formNames);
    } catch (getError) {
      console.error('Error getting rows for deletion:', getError);
      // Try using ZCQL as fallback
      try {
        const queryResult = await zcql.executeZCQLQuery(`SELECT ROWID, FormName FROM ${TABLE_NAME}`);
        allRows = queryResult.map(r => ({ 
          ROWID: r[TABLE_NAME].ROWID,
          FormName: r[TABLE_NAME].FormName 
        }));
        console.log(`Found ${allRows.length} records to delete (via ZCQL)`);
        
        // Extract formNames
        formNames = allRows
          .map(row => row.FormName)
          .filter(name => name && String(name).trim());
      } catch (zcqlError) {
        console.error('Error getting rows via ZCQL:', zcqlError);
        throw new Error(`Failed to retrieve records: ${getError.message}`);
      }
    }
    
    if (!allRows || allRows.length === 0) {
      return {
        status: 'success',
        message: 'No records found to delete',
        deletedCount: 0
      };
    }
    
    // Delete related Checklist entries first
    let checklistDeleteResult = { deletedCount: 0, errors: [] };
    if (formNames.length > 0) {
      checklistDeleteResult = await deleteChecklistEntriesByFormNames(catalyst, formNames);
      console.log(`Deleted ${checklistDeleteResult.deletedCount} related Checklist entries`);
    }
    
    // Delete all rows by their ROWID with better error handling
    let successCount = 0;
    let failCount = 0;
    const errors = [];
    
    for (const row of allRows) {
      try {
        const rowId = row.ROWID || row[TABLE_NAME]?.ROWID;
        if (!rowId) {
          console.warn('Row missing ROWID:', row);
          failCount++;
          continue;
        }
        await table.deleteRow(rowId);
        successCount++;
      } catch (deleteError) {
        console.error(`Error deleting row ${row.ROWID || 'unknown'}:`, deleteError);
        failCount++;
        errors.push({
          rowId: row.ROWID || row[TABLE_NAME]?.ROWID || 'unknown',
          error: deleteError.message
        });
      }
    }
    
    console.log(`Bulk delete completed. Success: ${successCount}, Failed: ${failCount}`);
    
    const message = `Successfully deleted ${successCount} checklist bulk records${checklistDeleteResult.deletedCount > 0 ? ` and ${checklistDeleteResult.deletedCount} related Checklist entries` : ''}`;
    
    if (failCount === 0) {
      return {
        status: 'success',
        message: message,
        deletedCount: successCount,
        checklistDeletedCount: checklistDeleteResult.deletedCount
      };
    } else if (successCount > 0) {
      return {
        status: 'partial',
        message: `Deleted ${successCount} records, but ${failCount} failed. ${checklistDeleteResult.deletedCount} Checklist entries deleted.`,
        deletedCount: successCount,
        failedCount: failCount,
        checklistDeletedCount: checklistDeleteResult.deletedCount,
        errors: errors
      };
    } else {
      return {
        status: 'error',
        message: `Failed to delete all ${allRows.length} records`,
        deletedCount: 0,
        failedCount: failCount,
        checklistDeletedCount: checklistDeleteResult.deletedCount,
        errors: errors
      };
    }
  } catch (error) {
    console.error('Error in bulk delete:', error);
    console.error('Error stack:', error.stack);
    return {
      status: 'error',
      message: 'Failed to bulk delete checklist data',
      error: error.message || error.toString(),
      errorDetails: {
        name: error.name,
        code: error.code,
        stack: error.stack
      }
    };
  }
};

// Bulk import checklist data
const bulkImportChecklistData = async (catalyst, dataArray) => {
  try {
    console.log(`Bulk importing ${dataArray.length} checklist bulk records`);
    console.log('Table name being used:', TABLE_NAME);
    console.log('Input data sample:', dataArray.slice(0, 2));
    
    const dataStore = catalyst.datastore();
    const table = dataStore.table(TABLE_NAME);
    
    // Check if table exists and get table info
    try {
      const tableInfo = await table.getTableInfo();
      console.log('Table info:', tableInfo);
    } catch (tableError) {
      console.error('Error getting table info:', tableError);
    }
    
    const formMasterLookup = await buildFormMasterFileLookup(catalyst);
    const records = dataArray.map((data) =>
      createChecklistBulkRecord(data, resolveFormFileForBulkRecord(data, formMasterLookup))
    );
    const linkedCount = records.filter((r) => r.FormFile && String(r.FormFile).trim()).length;
    console.log(`Form Master link: ${linkedCount}/${records.length} checklistbulk row(s) have FormFile`);
    
    console.log('Records to be inserted:', JSON.stringify(records.slice(0, 2), null, 2)); // Log first 2 records as sample
    console.log('Total records to insert:', records.length);
    
    // Log DueDate values to check if they're being preserved
    records.slice(0, 3).forEach((record, idx) => {
      console.log(`Record ${idx + 1} DueDate value:`, record.DueDate, `(type: ${typeof record.DueDate})`);
    });
    
    const result = await table.insertRows(records);
    
    console.log(`Bulk import completed. ${result.length} records added`);
    console.log('Result details:', JSON.stringify(result.slice(0, 2), null, 2)); // Log first 2 results as sample
    
    return {
      status: 'success',
      message: `Bulk import completed successfully. ${result.length} records added (${linkedCount} linked to Form Master templates)`,
      data: result.map((record, index) => ({
        id: record.ROWID,
        ...dataArray[index],
        formFile: records[index].FormFile || null,
        formFileName: resolveFormFileForBulkRecord(dataArray[index], formMasterLookup).formFileName || null
      }))
    };
  } catch (error) {
    console.error('Error in bulk import:', error);
    console.error('Error details:', error.message);
    console.error('Error stack:', error.stack);
    console.error('Error name:', error.name);
    console.error('Error code:', error.code);
    return {
      status: 'error',
      message: 'Failed to bulk import checklist data',
      error: error.message,
      errorDetails: {
        name: error.name,
        code: error.code,
        stack: error.stack
      }
    };
  }
};

// Get checklist data by sector
const getChecklistDataBySector = async (catalyst, sector) => {
  try {
    console.log(`Fetching checklist data for sector: ${sector}`);
    
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
      message: `Checklist data for sector '${sector}' retrieved successfully`,
      data: formattedData,
      count: formattedData.length
    };
  } catch (error) {
    console.error('Error fetching checklist data by sector:', error);
    return {
      status: 'error',
      message: 'Failed to fetch checklist data by sector',
      error: error.message
    };
  }
};

// Get checklist data by state
const getChecklistDataByState = async (catalyst, state) => {
  try {
    console.log(`Fetching checklist data for state: ${state}`);
    
    const dataStore = catalyst.datastore();
    const table = dataStore.table(TABLE_NAME);
    
    // Get all rows and filter by state
    const allRows = await table.getAllRows();
    const filteredRows = allRows.filter(record => 
      record.State && record.State.toLowerCase() === state.toLowerCase()
    );
    
    console.log(`Found ${filteredRows.length} records for state: ${state}`);
    
    const formattedData = filteredRows.map(record => convertToAppFormat(record));
    
    return {
      status: 'success',
      message: `Checklist data for state '${state}' retrieved successfully`,
      data: formattedData,
      count: formattedData.length
    };
  } catch (error) {
    console.error('Error fetching checklist data by state:', error);
    return {
      status: 'error',
      message: 'Failed to fetch checklist data by state',
      error: error.message
    };
  }
};

// Get checklist data by act
const getChecklistDataByAct = async (catalyst, act) => {
  try {
    console.log(`Fetching checklist data for act: ${act}`);
    
    const dataStore = catalyst.datastore();
    const table = dataStore.table(TABLE_NAME);
    
    // Get all rows and filter by act
    const allRows = await table.getAllRows();
    const filteredRows = allRows.filter(record => 
      record.Act && record.Act.toLowerCase().includes(act.toLowerCase())
    );
    
    console.log(`Found ${filteredRows.length} records for act: ${act}`);
    
    const formattedData = filteredRows.map(record => convertToAppFormat(record));
    
    return {
      status: 'success',
      message: `Checklist data for act '${act}' retrieved successfully`,
      data: formattedData,
      count: formattedData.length
    };
  } catch (error) {
    console.error('Error fetching checklist data by act:', error);
    return {
      status: 'error',
      message: 'Failed to fetch checklist data by act',
      error: error.message
    };
  }
};

/** Backfill FormFile / FormFileName on existing checklistbulk rows from Form Master. */
const syncChecklistBulkFormFiles = async (catalyst) => {
  try {
    const dataStore = catalyst.datastore();
    const table = dataStore.table(TABLE_NAME);
    const formMasterLookup = await buildFormMasterFileLookup(catalyst);
    if (formMasterLookup.size === 0) {
      return {
        status: 'success',
        message: 'No Form Master templates with uploaded files found. Upload templates in Form Master first.',
        updatedCount: 0,
        totalRows: 0
      };
    }

    const allRows = await table.getAllRows();
    let updatedCount = 0;
    let linkedCount = 0;
    const errors = [];

    for (const row of allRows || []) {
      const rowId = row.ROWID;
      if (!rowId) continue;
      const appRow = {
        formName: row.FormName,
        act: row.Act,
        description: row.Description,
        sector: row.Sector,
        state: row.State
      };
      const resolved = resolveFormFileForBulkRecord(appRow, formMasterLookup);
      const prevFile = String(row.FormFile || '').trim();
      if (!resolved.formFile) {
        const exactKey =
          formMasterLookup.get(buildChecklistBulkMatchKey(appRow)) ||
          formMasterLookup.get(buildChecklistBulkIdentityKey(appRow));
        if (prevFile && !exactKey?.formFile) {
          try {
            const patch = createChecklistBulkRecord(
              {
                sector: row.Sector,
                state: row.State,
                act: row.Act,
                formName: row.FormName,
                concernedGovtDepartment: row.ConcernedGovtDepartment,
                dueDate: row.DueDate,
                description: row.Description,
                nameOfTheCode: row.Nameofthecode,
                frequency: row.Frequency,
                nameOfTheRule: row.NameoftheRule
              },
              { formFile: '', formFileName: '' }
            );
            await table.updateRow(rowId, patch);
            updatedCount += 1;
          } catch (updateErr) {
            errors.push({ rowId, error: updateErr.message });
          }
        }
        continue;
      }
      linkedCount += 1;
      if (prevFile === String(resolved.formFile).trim()) {
        continue;
      }
      try {
        const patch = createChecklistBulkRecord(
          {
            sector: row.Sector,
            state: row.State,
            act: row.Act,
            formName: row.FormName,
            concernedGovtDepartment: row.ConcernedGovtDepartment,
            dueDate: row.DueDate,
            description: row.Description,
            nameOfTheCode: row.Nameofthecode,
            frequency: row.Frequency,
            nameOfTheRule: row.NameoftheRule
          },
          resolved
        );
        await table.updateRow(rowId, patch);
        updatedCount += 1;
      } catch (updateErr) {
        errors.push({ rowId, error: updateErr.message });
      }
    }

    return {
      status: errors.length > 0 && updatedCount === 0 ? 'error' : 'success',
      message: `Linked ${linkedCount} row(s) to Form Master; updated ${updatedCount} record(s) in checklistbulk.`,
      updatedCount,
      linkedCount,
      totalRows: (allRows || []).length,
      errors: errors.length > 0 ? errors : undefined
    };
  } catch (error) {
    console.error('Error syncing checklistbulk FormFile:', error);
    return {
      status: 'error',
      message: 'Failed to sync FormFile from Form Master',
      error: error.message
    };
  }
};

// Get count of checklist bulk records
const getChecklistBulkCount = async (catalyst) => {
  try {
    console.log('Getting checklist bulk data count...');
    
    const dataStore = catalyst.datastore();
    const table = dataStore.table(TABLE_NAME);
    
    // Use getAllRows and get length
    const result = await table.getAllRows();
    
    return {
      status: 'success',
      message: 'Checklist bulk data count retrieved successfully',
      count: result.length
    };
  } catch (error) {
    console.error('Error getting checklist bulk count:', error);
    return {
      status: 'error',
      message: 'Failed to get checklist bulk data count',
      error: error.message
    };
  }
};

// Populate table with sample data
const populateTableWithSampleData = async (catalyst) => {
  try {
    console.log('Populating Checklistbulk table with sample data...');
    
    const sampleData = [
      {
        Sector: "Manufacturing",
        State: "Maharashtra",
        Act: "Factories Act, 1948",
        FormName: "Form 3",
        ConcernedGovtDepartment: "Labour Department",
        DueDate: "31-Aug-2024",
        Description: "Annual return for manufacturing units",
        Nameofthecode: "FAC-AR-001",
        Frequency: "Annual",
        NameoftheRule: "Factories Rules"
      },
      {
        Sector: "Manufacturing",
        State: "Karnataka",
        Act: "Factories Act, 1948",
        FormName: "Form 4",
        ConcernedGovtDepartment: "Labour Department",
        DueDate: "31-Oct-2024",
        Description: "Quarterly return submission",
        Nameofthecode: "FAC-QR-004",
        Frequency: "Quarterly",
        NameoftheRule: "Factories Rules"
      },
      {
        Sector: "Healthcare",
        State: "Tamil Nadu",
        Act: "Clinical Establishments Act, 2010",
        FormName: "Form 3-A",
        ConcernedGovtDepartment: "Health Department",
        DueDate: "31-Oct-2024",
        Description: "Registration renewal for clinical establishments",
        Nameofthecode: "HEA-RN-003A",
        Frequency: "Annual",
        NameoftheRule: "Clinical Establishments Rules"
      },
      {
        Sector: "Finance",
        State: "Delhi",
        Act: "Banking Regulation Act, 1949",
        FormName: "Form 4(6)",
        ConcernedGovtDepartment: "RBI",
        DueDate: "15-Dec-2024",
        Description: "Compliance report submission",
        Nameofthecode: "FIN-CR-046",
        Frequency: "Half Yearly",
        NameoftheRule: "Banking Regulation Rules"
      },
      {
        Sector: "Environment",
        State: "All States",
        Act: "Environment Protection Act, 1986",
        FormName: "Form 10",
        ConcernedGovtDepartment: "Environment Department",
        DueDate: "30-Sep-2024",
        Description: "Environmental clearance application",
        Nameofthecode: "ENV-EC-010",
        Frequency: "Annual",
        NameoftheRule: "Environment Protection Rules"
      },
      {
        Sector: "Manufacturing",
        State: "Gujarat",
        Act: "Factories Act, 1948",
        FormName: "Form 12",
        ConcernedGovtDepartment: "Labour Department",
        DueDate: "30-Sep-2024",
        Description: "Safety report submission",
        Nameofthecode: "FAC-SR-012",
        Frequency: "Monthly",
        NameoftheRule: "Factories Rules"
      },
      {
        Sector: "Manufacturing",
        State: "West Bengal",
        Act: "Factories Act, 1948",
        FormName: "Form 15",
        ConcernedGovtDepartment: "Labour Department",
        DueDate: "30-Sep-2024",
        Description: "Annual compliance report",
        Nameofthecode: "FAC-AC-015",
        Frequency: "Annual",
        NameoftheRule: "Factories Rules"
      },
      {
        Sector: "Manufacturing",
        State: "All States",
        Act: "Factories Act, 1948",
        FormName: "Accident Report",
        ConcernedGovtDepartment: "Labour Department",
        DueDate: "30-Sep-2024",
        Description: "Accident incident reporting",
        Nameofthecode: "FAC-AI-AR",
        Frequency: "Event Based",
        NameoftheRule: "Factories Rules"
      }
    ];
    
    const dataStore = catalyst.datastore();
    const table = dataStore.table(TABLE_NAME);
    const result = await table.insertRows(sampleData);
    
    console.log(`Successfully populated table with ${result.length} sample records`);
    
    return {
      status: 'success',
      message: `Successfully populated Checklistbulk table with ${result.length} sample records`,
      data: result.map(record => ({
        id: record.ROWID,
        sector: record.Sector,
        state: record.State,
        act: record.Act,
        formName: record.FormName,
        concernedGovtDepartment: record.ConcernedGovtDepartment,
        dueDate: record.DueDate,
        description: record.Description,
        nameOfTheCode: record.Nameofthecode,
        frequency: record.Frequency,
        nameOfTheRule: record.NameoftheRule
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
  res.status(200).json({ status: 'success', message: 'checklistbulk_function ready' });
});

// Test endpoint to check table access and verify data
app.get('/test-table', async (req, res) => {
  try {
    console.log('Testing Checklistbulk table access...');
    const catalyst = res.locals.catalyst;
    const dataStore = catalyst.datastore();
    const table = dataStore.table(TABLE_NAME);
    
    console.log('Table name:', TABLE_NAME);
    console.log('Table object:', table);
    
    // Get table schema information
    let tableInfo = null;
    try {
      tableInfo = await table.getTableInfo();
      console.log('Table info:', tableInfo);
    } catch (schemaError) {
      console.error('Error getting table schema:', schemaError);
    }
    
    // Try to get all rows to see if table exists and has data
    const allRows = await table.getAllRows();
    console.log('Total rows in table:', allRows.length);
    console.log('Sample rows:', allRows.slice(0, 3));
    
    // Also try ZCQL query as alternative
    const zcql = catalyst.zcql();
    const countQuery = await zcql.executeZCQLQuery(`SELECT COUNT(ROWID) as count FROM ${TABLE_NAME}`);
    console.log('ZCQL count result:', countQuery);
    
    // Try to get table schema using ZCQL
    let schemaInfo = null;
    try {
      const schemaQuery = await zcql.executeZCQLQuery(`SELECT * FROM ${TABLE_NAME} LIMIT 1`);
      console.log('Schema query result:', schemaQuery);
      if (schemaQuery.length > 0) {
        schemaInfo = Object.keys(schemaQuery[0][TABLE_NAME] || {});
      }
    } catch (schemaError) {
      console.error('Error getting schema via ZCQL:', schemaError);
    }
    
    res.status(200).json({ 
      status: 'success', 
      message: 'Table access successful',
      tableInfo: {
        tableName: TABLE_NAME,
        recordCount: allRows.length,
        zcqlCount: countQuery[0]?.[TABLE_NAME]?.count || 0,
        sampleData: allRows.slice(0, 3),
        tableSchema: tableInfo,
        columnNames: schemaInfo
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
      state: "Test State",
      act: "Test Act",
      formName: "Test Form",
      concernedGovtDepartment: "Test Department",
      dueDate: "31-Dec-2024",
      description: "Test Description",
      nameOfTheCode: "TEST-001",
      frequency: "Monthly",
      nameOfTheRule: "Test Rules"
    };
    
    console.log('Test data to be added:', testData);
    
    const result = await addChecklistBulkData(catalyst, testData);
    
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

// Test endpoint for bulk import debugging
app.post('/test-bulk-import', async (req, res) => {
  try {
    console.log('Test bulk import endpoint called');
    const catalyst = res.locals.catalyst;
    
    const testBulkData = [
      {
        sector: "Manufacturing",
        state: "Maharashtra",
        act: "Factories Act, 1948",
        formName: "Form 3",
        concernedGovtDepartment: "Labour Department",
        dueDate: "31-Aug-2024",
        description: "Annual return for manufacturing units",
        nameOfTheCode: "FAC-AR-001",
        frequency: "Annual",
        nameOfTheRule: "Factories Rules"
      },
      {
        sector: "Healthcare",
        state: "Tamil Nadu",
        act: "Clinical Establishments Act, 2010",
        formName: "Form 3-A",
        concernedGovtDepartment: "Health Department",
        dueDate: "31-Oct-2024",
        description: "Registration renewal for clinical establishments",
        nameOfTheCode: "HEA-RN-003A",
        frequency: "Annual",
        nameOfTheRule: "Clinical Establishments Rules"
      }
    ];
    
    console.log('Test bulk data to be added:', testBulkData);
    
    const result = await bulkImportChecklistData(catalyst, testBulkData);
    
    res.status(result.status === 'success' ? 200 : 400).json(result);
    
  } catch (error) {
    console.error('Test bulk import error:', error);
    res.status(500).json({
      status: 'error',
      message: 'Test bulk import failed',
      error: error.message
    });
  }
});

// Get all checklist bulk data
app.get('/checklistbulk', async (req, res) => {
  try {
    const catalyst = res.locals.catalyst;
    const { action, id, sector, state, act } = req.query;
    
    console.log(`Checklist Bulk API Request - Action: ${action}, ID: ${id}, Sector: ${sector}, State: ${state}, Act: ${act}`);
    
    let result;
    
    switch (action) {
      case 'getAll':
        result = await getAllChecklistBulkData(catalyst);
        break;
        
      case 'getById':
        if (!id) {
          result = {
            status: 'error',
            message: 'ID parameter is required for getById action'
          };
        } else {
          result = await getChecklistBulkById(catalyst, id);
        }
        break;
        
      case 'getBySector':
        if (!sector) {
          result = {
            status: 'error',
            message: 'Sector parameter is required for getBySector action'
          };
        } else {
          result = await getChecklistDataBySector(catalyst, sector);
        }
        break;
        
      case 'getByState':
        if (!state) {
          result = {
            status: 'error',
            message: 'State parameter is required for getByState action'
          };
        } else {
          result = await getChecklistDataByState(catalyst, state);
        }
        break;
        
      case 'getByAct':
        if (!act) {
          result = {
            status: 'error',
            message: 'Act parameter is required for getByAct action'
          };
        } else {
          result = await getChecklistDataByAct(catalyst, act);
        }
        break;
        
      case 'count':
        result = await getChecklistBulkCount(catalyst);
        break;
        
      case 'populate':
        result = await populateTableWithSampleData(catalyst);
        break;

      case 'syncFormFiles':
        result = await syncChecklistBulkFormFiles(catalyst);
        break;
        
      default:
        result = {
          status: 'error',
          message: 'Invalid action. Supported actions: getAll, getById, getBySector, getByState, getByAct, count, populate, syncFormFiles'
        };
    }
    
    res.status(result.status === 'success' ? 200 : 400).json(result);
    
  } catch (error) {
    console.error('Checklist Bulk API Error:', error);
    res.status(500).json({
      status: 'error',
      message: 'Internal server error',
      error: error.message
    });
  }
});

// POST routes for data modification
app.post('/checklistbulk', async (req, res) => {
  try {
    const catalyst = res.locals.catalyst;
    const { action } = req.query;
    const body = req.body;
    
    console.log(`Checklist Bulk API POST Request - Action: ${action}`);
    
    let result;
    
    switch (action) {
      case 'add':
        if (!body) {
          result = {
            status: 'error',
            message: 'Request body is required for add action'
          };
        } else {
          result = await addChecklistBulkData(catalyst, body);
        }
        break;
        
      case 'bulkImport':
        if (!body || !Array.isArray(body)) {
          result = {
            status: 'error',
            message: 'Request body must be an array for bulkImport action'
          };
        } else {
          result = await bulkImportChecklistData(catalyst, body);
        }
        break;

      case 'syncFormFiles':
        result = await syncChecklistBulkFormFiles(catalyst);
        break;
        
      default:
        result = {
          status: 'error',
          message: 'Invalid actions: add, bulkImport, syncFormFiles'
        };
    }
    
    res.status(result.status === 'success' ? 200 : 400).json(result);
    
  } catch (error) {
    console.error('Checklist Bulk API Error:', error);
    res.status(500).json({
      status: 'error',
      message: 'Internal server error',
      error: error.message
    });
  }
});

// PUT route for updates
app.put('/checklistbulk', async (req, res) => {
  try {
    const catalyst = res.locals.catalyst;
    const { action, id } = req.query;
    const body = req.body;
    
    console.log(`Checklist Bulk API PUT Request - Action: ${action}, ID: ${id}`);
    
    let result;
    
    switch (action) {
      case 'update':
        if (!id || !body) {
          result = {
            status: 'error',
            message: 'ID and request body are required for update action'
          };
        } else {
          result = await updateChecklistBulkData(catalyst, id, body);
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
    console.error('Checklist Bulk API Error:', error);
    res.status(500).json({
      status: 'error',
      message: 'Internal server error',
      error: error.message
    });
  }
});

// DELETE route for deletions
app.delete('/checklistbulk', async (req, res) => {
  try {
    const catalyst = res.locals.catalyst;
    const { action, id } = req.query;
    
    console.log(`Checklist Bulk API DELETE Request - Action: ${action}, ID: ${id}`);
    
    let result;
    
    switch (action) {
      case 'delete':
        if (!id) {
          result = {
            status: 'error',
            message: 'ID parameter is required for delete action'
          };
        } else {
          result = await deleteChecklistBulkData(catalyst, id);
        }
        break;
        
      case 'bulkDelete':
        result = await bulkDeleteChecklistData(catalyst);
        break;
        
      default:
        result = {
          status: 'error',
          message: 'Invalid action. Supported actions: delete, bulkDelete'
        };
    }
    
    res.status(result.status === 'success' ? 200 : 400).json(result);
    
  } catch (error) {
    console.error('Checklist Bulk API Error:', error);
    res.status(500).json({
      status: 'error',
      message: 'Internal server error',
      error: error.message
    });
  }
});

module.exports = app;