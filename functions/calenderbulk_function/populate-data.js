const { ZCDataStore } = require('zcatalyst-sdk-node');

// Table name for calendar bulk data
const TABLE_NAME = 'Calenderbulk';

// Test data to populate the Calenderbulk table
const testData = [
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
  },
  {
    Forms: "Form 16",
    Sector: "Healthcare",
    Description: "Medical Examination Record",
    Rules: "Rule 16",
    ScheduleOfSubmission: "Medical examination of workers to be conducted annually",
    Jan: "15",
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
    December: ""
  },
  {
    Forms: "Form 18",
    Sector: "Finance",
    Description: "Annual Return",
    Rules: "Rule 18",
    ScheduleOfSubmission: "Annual return to be submitted by 31st March every year",
    Jan: "",
    Feb: "",
    Mar: "31",
    Apr: "",
    May: "",
    June: "",
    July: "",
    Aug: "",
    Sep: "",
    Oct: "",
    Nov: "",
    December: ""
  }
];

// Function to populate the Calenderbulk table
const populateCalenderbulkTable = async () => {
  try {
    console.log('Starting to populate Calenderbulk table...');
    
    const dataStore = ZCDataStore.getInstance();
    const table = dataStore.table(TABLE_NAME);
    
    // Insert all test data
    const result = await table.insertRows(testData);
    
    console.log(`Successfully inserted ${result.length} records into Calenderbulk table`);
    console.log('Inserted records:', result.map(record => ({
      ROWID: record.ROWID,
      Forms: record.Forms,
      Sector: record.Sector
    })));
    
    return {
      success: true,
      message: `Successfully populated Calenderbulk table with ${result.length} records`,
      records: result
    };
    
  } catch (error) {
    console.error('Error populating Calenderbulk table:', error);
    return {
      success: false,
      message: 'Failed to populate Calenderbulk table',
      error: error.message
    };
  }
};

// Function to verify data was inserted
const verifyDataInsertion = async () => {
  try {
    console.log('Verifying data insertion...');
    
    const dataStore = ZCDataStore.getInstance();
    const table = dataStore.table(TABLE_NAME);
    
    // Use getAllRows instead of select
    const result = await table.getAllRows();
    
    console.log(`Found ${result.length} records in Calenderbulk table`);
    
    if (result.length > 0) {
      console.log('Sample records:');
      result.slice(0, 3).forEach((record, index) => {
        console.log(`Record ${index + 1}:`, {
          ROWID: record.ROWID,
          Forms: record.Forms,
          Sector: record.Sector,
          Description: record.Description,
          Rules: record.Rules
        });
      });
    }
    
    return {
      success: true,
      count: result.length,
      records: result
    };
    
  } catch (error) {
    console.error('Error verifying data insertion:', error);
    return {
      success: false,
      error: error.message
    };
  }
};

// Main execution function
const main = async () => {
  console.log('=== Calenderbulk Table Population Script ===');
  
  // First, verify current state
  console.log('\n1. Checking current table state...');
  const currentState = await verifyDataInsertion();
  
  if (currentState.success && currentState.count > 0) {
    console.log(`Table already contains ${currentState.count} records.`);
    console.log('Do you want to add more records? (This script will add additional records)');
  }
  
  // Populate the table
  console.log('\n2. Populating table with test data...');
  const populateResult = await populateCalenderbulkTable();
  
  if (populateResult.success) {
    console.log('✅ ' + populateResult.message);
    
    // Verify the insertion
    console.log('\n3. Verifying data insertion...');
    const verifyResult = await verifyDataInsertion();
    
    if (verifyResult.success) {
      console.log(`✅ Verification successful. Table now contains ${verifyResult.count} records.`);
      console.log('\n🎉 Calenderbulk table has been successfully populated!');
      console.log('You can now check the Data Store Data View to see the records.');
    } else {
      console.log('❌ Verification failed:', verifyResult.error);
    }
  } else {
    console.log('❌ ' + populateResult.message);
    console.log('Error details:', populateResult.error);
  }
};

// Export functions for use in other modules
module.exports = {
  populateCalenderbulkTable,
  verifyDataInsertion,
  testData
};

// Run the script if called directly
if (require.main === module) {
  main().catch(console.error);
}
