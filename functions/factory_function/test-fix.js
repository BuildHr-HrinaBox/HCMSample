// Test script to verify Factory function is working with correct column mapping
const fetch = require('node-fetch');

const BASE_URL = 'http://localhost:3000'; // Adjust if your server runs on different port

async function testFactoryFunction() {
  console.log('Testing Factory function with corrected column mapping...\n');
  
  try {
    // Test 1: Health check
    console.log('1. Testing health check...');
    const healthResponse = await fetch(`${BASE_URL}/`);
    const healthResult = await healthResponse.json();
    console.log('Health check result:', healthResult);
    
    // Test 2: Test table access
    console.log('\n2. Testing table access...');
    const tableResponse = await fetch(`${BASE_URL}/test-table`);
    const tableResult = await tableResponse.json();
    console.log('Table test result:', tableResult);
    
    // Test 3: Test bulk import with correct data structure
    console.log('\n3. Testing bulk import with correct data structure...');
    const testData = [
      {
        formName: 'Test Form 1',
        factoryName: '', // Not in schema
        location: '', // Not in schema
        sector: '', // Not in schema
        pdfFile: null,
        hasPdf: false
      },
      {
        formName: 'Test Form 2',
        factoryName: '', // Not in schema
        location: '', // Not in schema
        sector: '', // Not in schema
        pdfFile: null,
        hasPdf: false
      }
    ];
    
    const bulkResponse = await fetch(`${BASE_URL}/factory?action=bulkImport`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
      },
      body: JSON.stringify(testData)
    });
    
    const bulkResult = await bulkResponse.json();
    console.log('Bulk import result:', bulkResult);
    
    // Test 4: Test getting all data
    console.log('\n4. Testing get all data...');
    const getAllResponse = await fetch(`${BASE_URL}/factory?action=getAll`);
    const getAllResult = await getAllResponse.json();
    console.log('Get all result:', getAllResult);
    
    console.log('\n✅ All tests completed successfully!');
    
  } catch (error) {
    console.error('❌ Test failed:', error.message);
  }
}

// Run the test
testFactoryFunction();
