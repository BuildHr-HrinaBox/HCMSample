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

// Health check
app.get('/', (req, res) => {
  res.status(200).json({ status: 'success', message: 'sitemanagement_function ready' });
});

function isMissingColumnError(err) {
  return /invalid|unknown|no such|column|does not exist/i.test(String(err?.message || err || ''));
}

function omitKeys(obj, keys) {
  const next = { ...obj };
  keys.forEach((k) => {
    delete next[k];
  });
  return next;
}

const OPTIONAL_SITE_COLUMNS = [
  'Audit',
  'Company',
  'ContractorName',
  'ContractorAddress',
  'ContractorEmail',
  'ContractorPhone',
  'ContractorCity',
  'ContractorState',
  'SandERCNumber',
  'FactoryRCNumber',
  'CLRARCNumber'
];

/**
 * Persist Site row with progressive fallbacks.
 * IMPORTANT: never drop `Company` when retrying for a missing `Audit` column —
 * that was wiping company name on every save.
 */
async function insertSiteWithFallback(table, insertData) {
  const attempts = [
    // Prefer saving Company; Audit is known-missing on many Site tables
    omitKeys(insertData, ['Audit']),
    insertData,
    omitKeys(insertData, ['Audit', 'SandERCNumber', 'FactoryRCNumber', 'CLRARCNumber']),
    omitKeys(insertData, [
      'Audit',
      'ContractorName',
      'ContractorAddress',
      'ContractorEmail',
      'ContractorPhone',
      'ContractorCity',
      'ContractorState',
      'SandERCNumber',
      'FactoryRCNumber',
      'CLRARCNumber'
    ]),
    // Last resort only — Company column itself missing
    omitKeys(insertData, OPTIONAL_SITE_COLUMNS)
  ];

  let lastErr;
  for (let i = 0; i < attempts.length; i += 1) {
    try {
      const payload = attempts[i];
      console.log(
        `Site insert attempt ${i + 1}/${attempts.length}; hasCompany=${Object.prototype.hasOwnProperty.call(payload, 'Company')}; company=${payload.Company || ''}`
      );
      return await table.insertRow(payload);
    } catch (err) {
      lastErr = err;
      if (!isMissingColumnError(err)) throw err;
      console.warn(`Site insert attempt ${i + 1} failed (missing column):`, err.message || err);
    }
  }
  throw lastErr;
}

async function updateSiteWithFallback(table, updateData) {
  const attempts = [
    omitKeys(updateData, ['Audit']),
    updateData,
    omitKeys(updateData, ['Audit', 'SandERCNumber', 'FactoryRCNumber', 'CLRARCNumber']),
    omitKeys(updateData, [
      'Audit',
      'ContractorName',
      'ContractorAddress',
      'ContractorEmail',
      'ContractorPhone',
      'ContractorCity',
      'ContractorState',
      'SandERCNumber',
      'FactoryRCNumber',
      'CLRARCNumber'
    ]),
    omitKeys(updateData, OPTIONAL_SITE_COLUMNS)
  ];

  let lastErr;
  for (let i = 0; i < attempts.length; i += 1) {
    try {
      const payload = attempts[i];
      console.log(
        `Site update attempt ${i + 1}/${attempts.length}; hasCompany=${Object.prototype.hasOwnProperty.call(payload, 'Company')}; company=${payload.Company || ''}`
      );
      return await table.updateRow(payload);
    } catch (err) {
      lastErr = err;
      if (!isMissingColumnError(err)) throw err;
      console.warn(`Site update attempt ${i + 1} failed (missing column):`, err.message || err);
    }
  }
  throw lastErr;
}

/** After insert, force-write Company if the row was created without it. */
async function ensureSiteCompanySaved(table, rowId, companyName) {
  const name = String(companyName || '').trim();
  if (!rowId || !name) return;
  try {
    const row = await table.getRow(rowId);
    const existing = String(row?.Company || '').trim();
    if (existing === name) return;
    console.log('Patching Site.Company after save:', { rowId, name, existing });
    await table.updateRow({ ROWID: rowId, Company: name });
  } catch (err) {
    console.warn('Could not patch Site.Company:', err.message || err);
  }
}

// Test endpoint to check Site table columns
app.get('/test-columns', async (req, res) => {
  try {
    console.log('Testing Site table column access...');
    const { catalyst } = res.locals;
    const zcql = catalyst.zcql();
    
    // Try different table names
    const tableNames = ['Site', 'site', 'SITE', 'SiteManagement', 'site_management'];
    
    for (const tableName of tableNames) {
      try {
        console.log(`Testing table: ${tableName}`);
        const columnQuery = await zcql.executeZCQLQuery(`SELECT * FROM ${tableName} LIMIT 1`);
        console.log(`Column test result for ${tableName}:`, columnQuery);
        
        if (columnQuery.length > 0) {
          const columns = Object.keys(columnQuery[0][tableName]);
          console.log(`Available columns in ${tableName}:`, columns);
          
          res.status(200).send({ 
            status: 'success', 
            message: `Column access successful for table: ${tableName}`,
            tableName: tableName,
            columns: columns
          });
          return;
        } else {
          console.log(`Table ${tableName} exists but has no data`);
        }
      } catch (err) {
        console.log(`Table ${tableName} doesn't exist or error:`, err.message);
      }
    }
    
    res.status(200).send({ 
      status: 'success', 
      message: 'No Site table found with any of the tested names',
      testedTables: tableNames
    });
  } catch (err) {
    console.error('Column test error:', err);
    res.status(500).send({ 
      status: 'failure', 
      message: err.message || 'Failed to access columns',
      error: err.toString()
    });
  }
});

// Test endpoint to try inserting with different column names
app.post('/test-insert', async (req, res) => {
  try {
    console.log('Testing Site table insert with different column names...');
    const { catalyst } = res.locals;
    const table = catalyst.datastore().table('Site');
    
    // Test with minimal data first
    const minimalData = {
      SiteName: 'Test Site'
    };
    
    try {
      console.log('Testing with minimal data (just SiteName)');
      const result = await table.insertRow(minimalData);
      console.log('Success with minimal data:', result);
      
      // Clean up test record
      await table.deleteRow(result.ROWID);
      
      res.status(200).send({ 
        status: 'success', 
        message: 'Minimal data insert successful - SiteName column works',
        workingColumns: ['SiteName']
      });
      return;
    } catch (err) {
      console.log('Failed with minimal data:', err.message);
    }
    
    // Try different column name variations
    const columnVariations = [
      { name: 'SiteName', value: 'Test Site' },
      { name: 'siteName', value: 'Test Site' },
      { name: 'site_name', value: 'Test Site' },
      { name: 'SITE_NAME', value: 'Test Site' },
      { name: 'Name', value: 'Test Site' },
      { name: 'name', value: 'Test Site' }
    ];
    
    for (const column of columnVariations) {
      try {
        console.log(`Testing column name: ${column.name}`);
        const testInsert = { [column.name]: column.value };
        const result = await table.insertRow(testInsert);
        console.log(`Success with column: ${column.name}`, result);
        
        // Clean up test record
        await table.deleteRow(result.ROWID);
        
        res.status(200).send({ 
          status: 'success', 
          message: `Working column name found: ${column.name}`,
          workingColumn: column.name
        });
        return;
      } catch (err) {
        console.log(`Failed with column: ${column.name}`, err.message);
      }
    }
    
    res.status(400).send({ 
      status: 'failure', 
      message: 'No working column names found - table might not exist or have different structure'
    });
  } catch (err) {
    console.error('Test insert error:', err);
    res.status(500).send({ 
      status: 'failure', 
      message: err.message || 'Failed to test insert',
      error: err.toString()
    });
  }
});

// Create Site
app.post('/sitemanagement', async (req, res) => {
  try {
    console.log('Received request body:', JSON.stringify(req.body, null, 2));
    
    const {
      siteName,
      companyId,
      companyName,
      company,
      siteAddress,
      siteCity,
      siteState,
      sitePostalCode,
      unitNo,
      contractorName,
      contractorAddress,
      contractorEmail,
      contractorPhone,
      contractorCity,
      contractorState,
      inchargeName,
      inchargePhone,
      inchargeEmail,
      inchargeDesignation,
      industry,
      sandERCNumber,
      factoryRCNumber,
      clraRCNumber,
      location,
      audit
    } = req.body;

    const companyNameToSave = String(companyName || company || '').trim();

    // Validate required fields
    if (!siteName || !String(siteName).trim()) {
      return res.status(400).send({ status: 'failure', message: 'Site Name is required.' });
    }
    if (!companyNameToSave) {
      return res.status(400).send({ status: 'failure', message: 'Company is required.' });
    }
    if (!siteAddress || !String(siteAddress).trim()) {
      return res.status(400).send({ status: 'failure', message: 'Site Address is required.' });
    }
    if (!siteCity || !String(siteCity).trim()) {
      return res.status(400).send({ status: 'failure', message: 'Site City is required.' });
    }
    if (!siteState || !String(siteState).trim()) {
      return res.status(400).send({ status: 'failure', message: 'Site State is required.' });
    }
    if (!sitePostalCode || !String(sitePostalCode).trim()) {
      return res.status(400).send({ status: 'failure', message: 'Site Postal Code is required.' });
    }
    if (!unitNo || !String(unitNo).trim()) {
      return res.status(400).send({ status: 'failure', message: 'Unit No is required.' });
    }
    if (!inchargeName || !String(inchargeName).trim()) {
      return res.status(400).send({ status: 'failure', message: 'Incharge Name is required.' });
    }
    if (!inchargePhone || !String(inchargePhone).trim()) {
      return res.status(400).send({ status: 'failure', message: 'Incharge Phone is required.' });
    }
    if (!inchargeDesignation || !String(inchargeDesignation).trim()) {
      return res.status(400).send({ status: 'failure', message: 'Incharge Designation is required.' });
    }

    const { catalyst } = res.locals;
    
    const table = catalyst.datastore().table('Site');
    
    // Insert all site data — Site table column is `Company` (company name)
    // Do not send Audit (column often missing); it previously caused fallbacks that dropped Company.
    const insertData = {
      SiteName: siteName,
      Company: companyNameToSave,
      SiteAddress: siteAddress,
      SiteCity: siteCity,
      SiteState: siteState,
      SitePostalCode: sitePostalCode,
      UNITNO: unitNo,
      ContractorName: contractorName || '',
      ContractorAddress: contractorAddress || '',
      ContractorEmail: contractorEmail || '',
      ContractorPhone: contractorPhone || '',
      ContractorCity: contractorCity || '',
      ContractorState: contractorState || '',
      InchargeName: inchargeName,
      InchargePhone: inchargePhone,
      InchargeEmail: inchargeEmail,
      InchargeDesignation: inchargeDesignation,
      Industry: industry,
      SandERCNumber: sandERCNumber || '',
      FactoryRCNumber: factoryRCNumber || '',
      CLRARCNumber: clraRCNumber || '',
      Location: location || ''
    };
    
    console.log('Inserting data to Site table:', JSON.stringify(insertData, null, 2));
    console.log('Company name to persist:', companyNameToSave, 'companyId:', companyId || '');
    
    const insertResp = await insertSiteWithFallback(table, insertData);
    console.log('Insert response:', JSON.stringify(insertResp, null, 2));

    await ensureSiteCompanySaved(table, insertResp.ROWID, companyNameToSave);
    
    const created = await table.getRow(insertResp.ROWID);
    console.log('Created record:', JSON.stringify(created, null, 2));
    
    // Convert to frontend format
    const siteDetail = {
      ROWID: created.ROWID,
      SiteName: created.SiteName || siteName,
      Company: created.Company || companyNameToSave,
      CompanyId: companyId || '',
      CompanyName: created.Company || companyNameToSave,
      SiteAddress: created.SiteAddress || siteAddress,
      SiteCity: created.SiteCity || siteCity,
      SiteState: created.SiteState || siteState,
      SitePostalCode: created.SitePostalCode || sitePostalCode,
      UnitNo: created.UNITNO || unitNo,
      ContractorName: created.ContractorName || contractorName || '',
      ContractorAddress: created.ContractorAddress || contractorAddress || '',
      ContractorEmail: created.ContractorEmail || contractorEmail || '',
      ContractorPhone: created.ContractorPhone || contractorPhone || '',
      ContractorCity: created.ContractorCity || contractorCity || '',
      ContractorState: created.ContractorState || contractorState || '',
      InchargeName: created.InchargeName || inchargeName,
      InchargePhone: created.InchargePhone || inchargePhone,
      InchargeEmail: created.InchargeEmail || inchargeEmail,
      InchargeDesignation: created.InchargeDesignation || inchargeDesignation,
      Industry: created.Industry || industry,
      SandERCNumber: created.SandERCNumber || sandERCNumber || '',
      FactoryRCNumber: created.FactoryRCNumber || factoryRCNumber || '',
      CLRARCNumber: created.CLRARCNumber || clraRCNumber || '',
      Location: created.Location || location || '',
      Audit: created.Audit || audit || 'false',
      CREATEDTIME: created.CREATEDTIME,
      MODIFIEDTIME: created.MODIFIEDTIME
    };
    
    res.status(200).send({ status: 'success', data: { siteDetail } });
  } catch (err) {
    console.error('Error creating site detail:', err);
    res.status(400).send({ status: 'failure', message: err.message || 'Failed to create site detail.' });
  }
});

// Update Site (ROWID from URL must match datastore row)
app.put('/sitemanagement/:ROWID', async (req, res) => {
  try {
    const { ROWID } = req.params;
    if (!ROWID || !String(ROWID).trim()) {
      return res.status(400).send({ status: 'failure', message: 'Site id (ROWID) is required.' });
    }

    const {
      siteName,
      companyId,
      companyName,
      company,
      siteAddress,
      siteCity,
      siteState,
      sitePostalCode,
      unitNo,
      contractorName,
      contractorAddress,
      contractorEmail,
      contractorPhone,
      contractorCity,
      contractorState,
      inchargeName,
      inchargePhone,
      inchargeEmail,
      inchargeDesignation,
      industry,
      sandERCNumber,
      factoryRCNumber,
      clraRCNumber,
      location,
      audit
    } = req.body;

    const companyNameToSave = String(companyName || company || '').trim();

    if (!siteName || !String(siteName).trim()) {
      return res.status(400).send({ status: 'failure', message: 'Site Name is required.' });
    }
    if (!companyNameToSave) {
      return res.status(400).send({ status: 'failure', message: 'Company is required.' });
    }
    if (!siteAddress || !String(siteAddress).trim()) {
      return res.status(400).send({ status: 'failure', message: 'Site Address is required.' });
    }
    if (!siteCity || !String(siteCity).trim()) {
      return res.status(400).send({ status: 'failure', message: 'Site City is required.' });
    }
    if (!siteState || !String(siteState).trim()) {
      return res.status(400).send({ status: 'failure', message: 'Site State is required.' });
    }
    if (!sitePostalCode || !String(sitePostalCode).trim()) {
      return res.status(400).send({ status: 'failure', message: 'Site Postal Code is required.' });
    }
    if (!unitNo || !String(unitNo).trim()) {
      return res.status(400).send({ status: 'failure', message: 'Unit No is required.' });
    }
    if (!inchargeName || !String(inchargeName).trim()) {
      return res.status(400).send({ status: 'failure', message: 'Incharge Name is required.' });
    }
    if (!inchargePhone || !String(inchargePhone).trim()) {
      return res.status(400).send({ status: 'failure', message: 'Incharge Phone is required.' });
    }
    if (!inchargeDesignation || !String(inchargeDesignation).trim()) {
      return res.status(400).send({ status: 'failure', message: 'Incharge Designation is required.' });
    }

    const { catalyst } = res.locals;
    const table = catalyst.datastore().table('Site');

    const updateData = {
      ROWID,
      SiteName: siteName,
      Company: companyNameToSave,
      SiteAddress: siteAddress,
      SiteCity: siteCity,
      SiteState: siteState,
      SitePostalCode: sitePostalCode,
      UNITNO: unitNo,
      ContractorName: contractorName || '',
      ContractorAddress: contractorAddress || '',
      ContractorEmail: contractorEmail || '',
      ContractorPhone: contractorPhone || '',
      ContractorCity: contractorCity || '',
      ContractorState: contractorState || '',
      InchargeName: inchargeName,
      InchargePhone: inchargePhone,
      InchargeEmail: inchargeEmail,
      InchargeDesignation: inchargeDesignation,
      Industry: industry,
      SandERCNumber: sandERCNumber || '',
      FactoryRCNumber: factoryRCNumber || '',
      CLRARCNumber: clraRCNumber || '',
      Location: location || ''
    };

    console.log('Updating Site row:', ROWID, JSON.stringify(updateData, null, 2));
    console.log('Company name to persist:', companyNameToSave, 'companyId:', companyId || '');
    await updateSiteWithFallback(table, updateData);
    await ensureSiteCompanySaved(table, ROWID, companyNameToSave);
    const updated = await table.getRow(ROWID);

    const siteDetail = {
      ROWID: updated.ROWID,
      SiteName: updated.SiteName || siteName,
      Company: updated.Company || companyNameToSave,
      CompanyId: companyId || '',
      CompanyName: updated.Company || companyNameToSave,
      SiteAddress: updated.SiteAddress || siteAddress,
      SiteCity: updated.SiteCity || siteCity,
      SiteState: updated.SiteState || siteState,
      SitePostalCode: updated.SitePostalCode || sitePostalCode,
      UnitNo: updated.UNITNO || unitNo,
      ContractorName: updated.ContractorName || contractorName || '',
      ContractorAddress: updated.ContractorAddress || contractorAddress || '',
      ContractorEmail: updated.ContractorEmail || contractorEmail || '',
      ContractorPhone: updated.ContractorPhone || contractorPhone || '',
      ContractorCity: updated.ContractorCity || contractorCity || '',
      ContractorState: updated.ContractorState || contractorState || '',
      InchargeName: updated.InchargeName || inchargeName,
      InchargePhone: updated.InchargePhone || inchargePhone,
      InchargeEmail: updated.InchargeEmail || inchargeEmail,
      InchargeDesignation: updated.InchargeDesignation || inchargeDesignation,
      Industry: updated.Industry || industry,
      SandERCNumber: updated.SandERCNumber || sandERCNumber || '',
      FactoryRCNumber: updated.FactoryRCNumber || factoryRCNumber || '',
      CLRARCNumber: updated.CLRARCNumber || clraRCNumber || '',
      Location: updated.Location || location || '',
      Audit: updated.Audit || audit || 'false',
      CREATEDTIME: updated.CREATEDTIME,
      MODIFIEDTIME: updated.MODIFIEDTIME
    };

    res.status(200).send({ status: 'success', data: { siteDetail } });
  } catch (err) {
    console.error('Error updating site detail:', err);
    res.status(400).send({ status: 'failure', message: err.message || 'Failed to update site detail.' });
  }
});

// List Sites with optional pagination
app.get('/sitemanagement', async (req, res) => {
  try {
    console.log('Fetching site details...');
    const { catalyst } = res.locals;
    const zcql = catalyst.zcql();
    const page = parseInt(req.query.page, 10) || 1;
    const perPage = parseInt(req.query.perPage, 10) || 50;
    const returnAll = !req.query.page && !req.query.perPage;
    const limitClause = returnAll ? '' : `LIMIT ${(page - 1) * perPage + 1},${perPage}`;
    
    console.log('Executing count query...');
    const countRows = await zcql.executeZCQLQuery('SELECT COUNT(ROWID) as count FROM Site');
    const countCell = countRows[0]?.Site || countRows[0]?.site || {};
    const total = parseInt(
      countCell.count ?? countCell.COUNT ?? countCell['COUNT(ROWID)'] ?? 0,
      10
    ) || 0;
    console.log('Total records:', total);
    
    console.log('Executing data query...');
    const siteSelectWithCompany =
      'ROWID, SiteName, Company, SiteAddress, SiteCity, SiteState, SitePostalCode, UNITNO, ContractorName, ContractorAddress, ContractorEmail, ContractorPhone, ContractorCity, ContractorState, InchargeName, InchargePhone, InchargeEmail, InchargeDesignation, Industry, SandERCNumber, FactoryRCNumber, CLRARCNumber, Location, CREATEDTIME, MODIFIEDTIME';
    const siteSelectFull =
      'ROWID, SiteName, SiteAddress, SiteCity, SiteState, SitePostalCode, UNITNO, ContractorName, ContractorAddress, ContractorEmail, ContractorPhone, ContractorCity, ContractorState, InchargeName, InchargePhone, InchargeEmail, InchargeDesignation, Industry, SandERCNumber, FactoryRCNumber, CLRARCNumber, Location, CREATEDTIME, MODIFIEDTIME';
    const siteSelectBase =
      'ROWID, SiteName, SiteAddress, SiteCity, SiteState, SitePostalCode, UNITNO, InchargeName, InchargePhone, InchargeEmail, InchargeDesignation, Industry, Location, CREATEDTIME, MODIFIEDTIME';
    const siteSelectWithRc =
      'ROWID, SiteName, SiteAddress, SiteCity, SiteState, SitePostalCode, UNITNO, InchargeName, InchargePhone, InchargeEmail, InchargeDesignation, Industry, SandERCNumber, FactoryRCNumber, CLRARCNumber, Location, CREATEDTIME, MODIFIEDTIME';
    let rows;
    let hasContractorColumns = true;
    let hasRcNumberColumns = true;
    let hasCompanyColumns = true;
    try {
      rows = await zcql.executeZCQLQuery(
        `SELECT ${siteSelectWithCompany} FROM Site ORDER BY ROWID DESC ${limitClause}`
      );
    } catch (companyErr) {
      const companyErrMsg = String(companyErr?.message || companyErr || '');
      if (/invalid|unknown|no such|column/i.test(companyErrMsg)) {
        console.warn('Site list: Company columns missing, trying without Company columns:', companyErrMsg);
        hasCompanyColumns = false;
        try {
          rows = await zcql.executeZCQLQuery(
            `SELECT ${siteSelectFull} FROM Site ORDER BY ROWID DESC ${limitClause}`
          );
        } catch (queryErr) {
          const errMsg = String(queryErr?.message || queryErr || '');
          if (/invalid|unknown|no such|column/i.test(errMsg)) {
            console.warn('Site list: contractor columns missing, trying without contractor columns:', errMsg);
            hasContractorColumns = false;
            try {
              rows = await zcql.executeZCQLQuery(
                `SELECT ${siteSelectWithRc} FROM Site ORDER BY ROWID DESC ${limitClause}`
              );
            } catch (rcErr) {
              const rcErrMsg = String(rcErr?.message || rcErr || '');
              if (/invalid|unknown|no such|column/i.test(rcErrMsg)) {
                console.warn('Site list: RC number columns missing, using base SELECT:', rcErrMsg);
                hasRcNumberColumns = false;
                rows = await zcql.executeZCQLQuery(
                  `SELECT ${siteSelectBase} FROM Site ORDER BY ROWID DESC ${limitClause}`
                );
              } else {
                throw rcErr;
              }
            }
          } else {
            throw queryErr;
          }
        }
      } else {
        throw companyErr;
      }
    }
    console.log('Raw rows from database:', JSON.stringify(rows, null, 2));
    if (rows?.[0]?.Site) {
      console.log('Site row keys:', Object.keys(rows[0].Site));
      console.log('Site.Company sample:', rows[0].Site.Company, rows[0].Site.company);
    }
    
    const siteDetails = rows.map(r => {
      const site = r.Site || {};
      const companyValue = hasCompanyColumns
        ? String(site.Company ?? site.company ?? site.CompanyName ?? site.companyName ?? '').trim()
        : '';
      return {
      id: site.ROWID,
      siteName: site.SiteName,
      companyId: String(site.CompanyId ?? site.companyId ?? '').trim(),
      companyName: companyValue,
      company: companyValue,
      siteAddress: site.SiteAddress,
      siteCity: site.SiteCity,
      siteState: site.SiteState,
      sitePostalCode: site.SitePostalCode,
      unitNo: site.UNITNO,
      contractorName: hasContractorColumns ? site.ContractorName : '',
      contractorAddress: hasContractorColumns ? site.ContractorAddress : '',
      contractorEmail: hasContractorColumns ? site.ContractorEmail : '',
      contractorPhone: hasContractorColumns ? site.ContractorPhone : '',
      contractorCity: hasContractorColumns ? site.ContractorCity : '',
      contractorState: hasContractorColumns ? site.ContractorState : '',
      inchargeName: site.InchargeName,
      inchargePhone: site.InchargePhone,
      inchargeEmail: site.InchargeEmail,
      inchargeDesignation: site.InchargeDesignation,
      industry: site.Industry,
      sandERCNumber: hasRcNumberColumns ? site.SandERCNumber : '',
      factoryRCNumber: hasRcNumberColumns ? site.FactoryRCNumber : '',
      clraRCNumber: hasRcNumberColumns ? site.CLRARCNumber : '',
      location: site.Location,
      audit: false, // Temporarily set to false since Audit column is commented out
      createdTime: site.CREATEDTIME,
      modifiedTime: site.MODIFIEDTIME
    };
    });
    
    console.log('Processed site details:', JSON.stringify(siteDetails, null, 2));
    res.status(200).send({ status: 'success', data: { siteDetails, total, hasMore: returnAll ? false : page * perPage < total } });
  } catch (err) {
    console.error('Error fetching site details:', err);
    res.status(500).send({ status: 'failure', message: err.message || 'Failed to fetch site details.' });
  }
});

// Delete Site
app.delete('/sitemanagement/:ROWID', async (req, res) => {
  try {
    const { ROWID } = req.params;
    const { catalyst } = res.locals;
    const table = catalyst.datastore().table('Site');
    await table.deleteRow(ROWID);
    res.status(200).send({ status: 'success', data: { id: ROWID } });
  } catch (err) {
    console.error('Error deleting site:', err);
    res.status(500).send({ status: 'failure', message: err.message || 'Failed to delete site.' });
  }
});

module.exports = app;
