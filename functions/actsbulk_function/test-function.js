const express = require('express');
const catalystSDK = require('zcatalyst-sdk-node');

// Create a simple test to check if the function works
async function testActsBulkFunction() {
  console.log('Testing Actsbulk Function...');
  
  try {
    // Simulate a request object (you'll need to replace this with actual request)
    const mockReq = {
      url: '/test-table',
      method: 'GET',
      headers: {},
      body: {}
    };
    
    const mockRes = {
      status: (code) => ({
        json: (data) => {
          console.log('Response Status:', code);
          console.log('Response Data:', JSON.stringify(data, null, 2));
        }
      }),
      locals: {}
    };
    
    // Initialize catalyst (this might fail in local testing)
    try {
      const catalyst = catalystSDK.initialize(mockReq);
      console.log('Catalyst initialized successfully');
      
      // Test table access
      const dataStore = catalyst.datastore();
      const table = dataStore.table('Actsbulk');
      console.log('Table object created:', !!table);
      
      // Try to get all rows
      const allRows = await table.getAllRows();
      console.log('Current rows in table:', allRows.length);
      console.log('Sample data:', allRows.slice(0, 2));
      
    } catch (catalystError) {
      console.log('Catalyst initialization failed (expected in local testing):', catalystError.message);
      console.log('This is normal when testing locally. The function needs to be deployed to Catalyst.');
    }
    
  } catch (error) {
    console.error('Test failed:', error);
  }
}

// Instructions for testing
console.log('=== ACTSBULK FUNCTION TEST ===');
console.log('');
console.log('This test script helps debug the Actsbulk function.');
console.log('');
console.log('STEP 1: Deploy the function to Catalyst');
console.log('STEP 2: Test the endpoints using these URLs:');
console.log('');
console.log('Test table access:');
console.log('GET /test-table');
console.log('');
console.log('Test adding a single record:');
console.log('POST /test-add');
console.log('');
console.log('Test getting all data:');
console.log('GET /actsbulk?action=getAll');
console.log('');
console.log('Test populating with sample data:');
console.log('GET /actsbulk?action=populate');
console.log('');
console.log('STEP 3: Check the logs in Catalyst console for detailed information');
console.log('');
console.log('=== RUNNING LOCAL TEST ===');

testActsBulkFunction();
