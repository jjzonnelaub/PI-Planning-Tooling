// ===== JIRA JQL DEBUGGER =====
// Run this to figure out why your query returns 0 results

/**
 * Debug function to test your JQL queries step by step
 * This will help identify which part of the query is failing
 */
function debugJQLQuery() {
  const ui = SpreadsheetApp.getUi();

  // Test queries progressively to narrow down the issue
  const tests = [
    {
      name: "Test 1: Basic Epic search",
      jql: "issuetype = Epic",
      description: "Just get any epics"
    },
    {
      name: "Test 2: Epic with PI field",
      jql: 'issuetype = Epic AND cf[10113] is not EMPTY',
      description: "Epics that have a PI field set"
    },
    {
      name: "Test 3: Epic with PI 13",
      jql: 'issuetype = Epic AND cf[10113] = "PI 13"',
      description: "Epics for PI 13"
    },
    {
      name: "Test 4: Epic with PI 13, not closed",
      jql: 'issuetype = Epic AND cf[10113] = "PI 13" AND status != "Closed"',
      description: "PI 13 epics that aren't closed"
    },
    {
      name: "Test 5: Epic with value stream",
      jql: 'issuetype = Epic AND cf[10046] is not EMPTY',
      description: "Epics with a value stream"
    },
    {
      name: "Test 6: Full EMA Clinical query",
      jql: 'issuetype = Epic AND cf[10113] = "PI 13" AND status != "Closed" AND cf[10046] = "EMA Clinical"',
      description: "Your full query for EMA Clinical"
    }
  ];

  console.log('=== STARTING JQL DEBUG TESTS ===\n');

  const results = [];

  tests.forEach(test => {
    console.log(`\n${test.name}`);
    console.log(`Description: ${test.description}`);
    console.log(`JQL: ${test.jql}`);

    try {
      const count = testJQL(test.jql);
      console.log(`✓ Result: ${count} issues found`);

      results.push({
        test: test.name,
        jql: test.jql,
        count: count,
        status: 'SUCCESS'
      });

      // If we get results, show some sample keys
      if (count > 0) {
        const samples = getSampleIssues(test.jql, 3);
        console.log(`Sample keys: ${samples.join(', ')}`);
      }

    } catch (error) {
      console.log(`✗ Error: ${error.message}`);
      results.push({
        test: test.name,
        jql: test.jql,
        count: 0,
        status: 'ERROR: ' + error.message
      });
    }
  });

  // Create summary
  console.log('\n\n=== SUMMARY ===');
  results.forEach(r => {
    console.log(`${r.test}: ${r.count} results (${r.status})`);
  });

  // Show in dialog
  let message = 'JQL Debug Results:\n\n';
  results.forEach(r => {
    message += `${r.test}\n`;
    message += `  Results: ${r.count}\n`;
    message += `  Status: ${r.status}\n\n`;
  });

  ui.alert('JQL Debug Complete', message, ui.ButtonSet.OK);
}

/**
 * Test a JQL query and return count
 */
function testJQL(jql) {
  const url = `${JIRA_CONFIG.baseUrl}/rest/api/3/search/jql`;

  const payload = {
    jql: jql,
    maxResults: 1,
    fields: ['key'],
    fieldsByKeys: false
  };

  const response = makeJiraRequest(url, 'POST', payload);
  return response.total || 0;
}

/**
 * Get sample issue keys for a JQL query
 */
function getSampleIssues(jql, maxSamples = 5) {
  const url = `${JIRA_CONFIG.baseUrl}/rest/api/3/search/jql`;

  const payload = {
    jql: jql,
    maxResults: maxSamples,
    fields: ['key'],
    fieldsByKeys: false
  };

  const response = makeJiraRequest(url, 'POST', payload);

  if (response && response.issues) {
    return response.issues.map(issue => issue.key);
  }

  return [];
}

/**
 * Debug specific field values
 * This checks what the actual field values are in JIRA
 */
function debugFieldValues() {
  console.log('=== DEBUGGING FIELD VALUES ===\n');

  // Get a sample epic
  const sampleJql = 'issuetype = Epic ORDER BY created DESC';
  const url = `${JIRA_CONFIG.baseUrl}/rest/api/3/search/jql`;

  const payload = {
    jql: sampleJql,
    maxResults: 5,
    fields: ['key', 'summary', 'customfield_10113', 'customfield_10046', 'status'],
    fieldsByKeys: false
  };

  try {
    const response = makeJiraRequest(url, 'POST', payload);

    if (response && response.issues && response.issues.length > 0) {
      console.log(`Found ${response.issues.length} sample epics:\n`);

      response.issues.forEach(issue => {
        console.log(`\n--- ${issue.key}: ${issue.fields.summary} ---`);
        console.log(`Status: ${issue.fields.status.name}`);

        const piField = issue.fields.customfield_10113;
        console.log(`PI Field (cf[10113]): ${JSON.stringify(piField)}`);

        const vsField = issue.fields.customfield_10046;
        console.log(`Value Stream (cf[10046]): ${JSON.stringify(vsField)}`);
      });

      // Now test what works
      console.log('\n\n=== TESTING DIFFERENT PI VALUE FORMATS ===');

      const piFormats = [
        '"PI 13"',
        '13',
        '"13"',
        'PI13',
        '"PI13"'
      ];

      piFormats.forEach(format => {
        const testJql = `issuetype = Epic AND cf[10113] = ${format}`;
        try {
          const count = testJQL(testJql);
          console.log(`cf[10113] = ${format} → ${count} results`);
        } catch (error) {
          console.log(`cf[10113] = ${format} → ERROR`);
        }
      });

    } else {
      console.log('No epics found at all!');
    }

  } catch (error) {
    console.error('Error debugging fields:', error);
  }
}

/**
 * Test what the correct PI field format should be
 */
function findCorrectPIFormat() {
  console.log('=== FINDING CORRECT PI FORMAT ===\n');

  // Get an epic that you know should be in PI 13
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    'Enter a JIRA Epic Key',
    'Enter the key of an epic you know is in PI 13 (e.g., ABC-123):',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  const epicKey = response.getResponseText().trim();

  if (!epicKey) {
    ui.alert('No epic key provided');
    return;
  }

  // Fetch this specific epic
  const url = `${JIRA_CONFIG.baseUrl}/rest/api/3/issue/${epicKey}`;

  try {
    const issue = makeJiraRequest(url, 'GET');

    console.log(`\n=== Epic ${epicKey} Field Values ===`);
    console.log(`Summary: ${issue.fields.summary}`);
    console.log(`Status: ${issue.fields.status.name}`);

    // Check PI field
    const piField = issue.fields.customfield_10113;
    console.log(`\nPI Field (customfield_10113):`);
    console.log(JSON.stringify(piField, null, 2));

    // Check Value Stream field
    const vsField = issue.fields.customfield_10046;
    console.log(`\nValue Stream Field (customfield_10046):`);
    console.log(JSON.stringify(vsField, null, 2));

    // Check status
    console.log(`\nStatus:`);
    console.log(JSON.stringify(issue.fields.status, null, 2));

    // Now test queries
    console.log('\n=== TESTING QUERIES WITH THIS EPIC ===');

    // Test 1: Just the key
    let testJql = `key = "${epicKey}"`;
    console.log(`\nTest: ${testJql}`);
    console.log(`Result: ${testJQL(testJql)} issues`);

    // Test 2: Key + issue type
    testJql = `key = "${epicKey}" AND issuetype = Epic`;
    console.log(`\nTest: ${testJql}`);
    console.log(`Result: ${testJQL(testJql)} issues`);

    // Test 3: Try with PI field if present
    if (piField) {
      let piValue;
      if (typeof piField === 'string') {
        piValue = piField;
      } else if (piField.value) {
        piValue = piField.value;
      } else if (piField.name) {
        piValue = piField.name;
      } else {
        piValue = JSON.stringify(piField);
      }

      testJql = `issuetype = Epic AND cf[10113] = "${piValue}"`;
      console.log(`\nTest: ${testJql}`);
      console.log(`Result: ${testJQL(testJql)} issues`);
    }

    // Show in UI
    ui.alert(
      'Field Values Found',
      `Check the Execution Logs for detailed field values from epic ${epicKey}`,
      ui.ButtonSet.OK
    );

  } catch (error) {
    console.error('Error fetching epic:', error);
    ui.alert('Error', `Could not fetch epic ${epicKey}: ${error.message}`, ui.ButtonSet.OK);
  }
}

/**
 * Quick test of your actual query
 */
function testMyActualQuery() {
  const ui = SpreadsheetApp.getUi();

  // Your actual query from the logs
  const jql = 'issuetype = Epic AND cf[10113] = "PI 13" AND status != "Closed" AND cf[10046] = "EMA Clinical"';

  console.log('Testing your actual query:');
  console.log(jql);

  try {
    const count = testJQL(jql);
    console.log(`Result: ${count} issues`);

    if (count > 0) {
      const samples = getSampleIssues(jql, 5);
      console.log('Sample epics found:', samples.join(', '));

      ui.alert('Success!', `Found ${count} epics:\n${samples.join('\n')}`, ui.ButtonSet.OK);
    } else {
      ui.alert('No Results', 'Query returned 0 results. Run debugJQLQuery() for detailed analysis.', ui.ButtonSet.OK);
    }

  } catch (error) {
    console.error('Error:', error);
    ui.alert('Error', error.message, ui.ButtonSet.OK);
  }
}

// ===== MENU INTEGRATION =====
// Add these functions to your menu so you can run them easily
function addDebugMenu() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🔍 Debug JIRA')
    .addItem('1. Test JQL Queries', 'debugJQLQuery')
    .addItem('2. Check Field Values', 'debugFieldValues')
    .addItem('3. Find Correct PI Format', 'findCorrectPIFormat')
    .addItem('4. Test Actual Query', 'testMyActualQuery')
    .addToUi();
}
function debugFieldParsingStandalone() {
  console.log('Starting field parsing debug...');

  try {
    const jql = 'issuetype = Epic AND status != "Closed" ORDER BY created DESC';

    // CORRECT: Use v3 endpoint
    const url = `${JIRA_CONFIG.baseUrl}/rest/api/3/search/jql`;

    // CORRECT: Send as POST payload
    const payload = {
      jql: jql,
      maxResults: 1
    };

    // CORRECT: Use POST method
    const response = makeJiraRequest(url, 'POST', payload);

    if (response && response.issues && response.issues.length > 0) {
      const issue = response.issues[0];
      console.log('=' .repeat(50));
      console.log('Testing field parsing on issue:', issue.key);
      console.log('=' .repeat(50));

      // Test each field mapping
      Object.entries(FIELD_MAPPINGS).forEach(([fieldName, fieldId]) => {
        const rawValue = issue.fields[fieldId];

        console.log(`\n${fieldName} (${fieldId}):`);
        console.log('  Raw value:', JSON.stringify(rawValue, null, 2));
      });

      const parsed = parseJiraIssue(issue);
      console.log('Full parsed issue:', JSON.stringify(parsed, null, 2));

    } else {
      console.log('No issues found in JIRA');
    }

  } catch (error) {
    console.error('Debug error:', error);
    console.error('Stack trace:', error.stack);
  }
}
function testTimeKeepersSummaryGeneration() {
  // Get the existing Time-Keepers summary sheet
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const summarySheet = spreadsheet.getSheetByName('PI 12 - Time-Keepers Summary');

  if (!summarySheet) {
    console.log('No Time-Keepers summary sheet found');
    return;
  }

  // Check what's in the sheet
  const values = summarySheet.getDataRange().getValues();
  console.log(`Summary sheet has ${values.length} rows`);

  // Check first 10 rows
  console.log('First 10 rows:');
  for (let i = 0; i < Math.min(10, values.length); i++) {
    console.log(`Row ${i + 1}: ${values[i][0]}`);
  }
}
function debugTimeKeepersIssues() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const piSheet = spreadsheet.getSheetByName('PI 12');

  if (!piSheet) {
    console.log('No PI 12 sheet found');
    return;
  }

  // Parse the data
  const values = piSheet.getDataRange().getValues();
  const headers = values[3];
  const issues = parsePISheetData(values, headers);

  // Filter for Time-Keepers
  const tkIssues = issues.filter(i => i.scrumTeam === 'Time-Keepers');
  const tkEpics = tkIssues.filter(i => i.issueType === 'Epic');
  const tkStories = tkIssues.filter(i => i.issueType !== 'Epic');

  console.log('Time-Keepers parsed data:');
  console.log(`- Total issues: ${tkIssues.length}`);
  console.log(`- Epics: ${tkEpics.length}`);
  console.log(`- Stories/Tasks: ${tkStories.length}`);

  if (tkEpics.length > 0) {
    console.log('\nFirst epic:');
    console.log(JSON.stringify(tkEpics[0], null, 2));
  }

  if (tkStories.length > 0) {
    console.log('\nFirst story:');
    console.log(JSON.stringify(tkStories[0], null, 2));
  }

  // Check capacity
  const capacityData = getCapacityDataForTeam(spreadsheet, 'Time-Keepers');
  console.log('\nCapacity data:', capacityData);
}

function testTimeKeepersProgressGauges() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const piSheet = spreadsheet.getSheetByName('PI 12');

  if (!piSheet) return;

  // Get Time-Keepers data
  const values = piSheet.getDataRange().getValues();
  const headers = values[3];
  const allIssues = parsePISheetData(values, headers);
  const tkIssues = allIssues.filter(i => i.scrumTeam === 'Time-Keepers');
  const tkEpics = tkIssues.filter(i => i.issueType === 'Epic');
  const tkStories = tkIssues.filter(i => i.issueType !== 'Epic');

  console.log('Testing planning progress calculation:');
  console.log(`- Issues: ${tkIssues.length}`);
  console.log(`- Epics: ${tkEpics.length}`);
  console.log(`- Stories: ${tkStories.length}`);

  // Test capacity lookup
  const capacityData = getCapacityDataForTeam(spreadsheet, 'Time-Keepers');
  console.log('- Capacity data:', capacityData);

  // Check for story points
  const storyPoints = tkStories.reduce((sum, s) => sum + (s.storyPoints || 0), 0);
  console.log(`- Total story points: ${storyPoints}`);

  // Check iteration slotting data
  const slottedData = calculateSlottedData(tkIssues, 12, 'Time-Keepers');
  console.log('- Slotted data:', JSON.stringify(slottedData, null, 2));
}

function createSafeIterationSlottingChart(sheet, startRow, issues, scrumTeam, programIncrement) {
  console.log(`Creating Iteration Slotting chart for ${scrumTeam} (SAFE VERSION)`);

  try {
    // Extract PI number
    const piNumber = parseInt(programIncrement.replace('PI ', ''));
    if (isNaN(piNumber)) {
      console.error('Invalid PI number in programIncrement:', programIncrement);
      return startRow;
    }

    // Title
    sheet.getRange(startRow, 1).setValue('Iteration Slotting');
    sheet.getRange(startRow, 1, 1, 11).setBackground('#E1D5E7');
    sheet.getRange(startRow, 1).setFontSize(14).setFontWeight('bold').setFontColor('black');
    sheet.getRange(startRow, 1).setFontFamily('Comfortaa');
    startRow++;

    // Headers
    const headers = [
      'Iteration', 'Baseline Capacity', 'Product Load', 'Slotted Product Load', 'Remaining',
      'Tech/Platform Load', 'Slotted Tech/Platform Load', 'Remaining',
      'Planned Quality Load', 'Slotted Planned Quality Load', 'Remaining'
    ];

    sheet.getRange(startRow, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(startRow, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#9b7bb8')
      .setFontColor('white')
      .setFontSize(8)
      .setWrap(true)
      .setFontFamily('Comfortaa');
    startRow++;

    // Calculate slotted values
    const slottedData = calculateSlottedData(issues, piNumber, scrumTeam);

    // Check if MMPM sheet exists
    const spreadsheet = sheet.getParent();
    const mmpmSheet = spreadsheet.getSheetByName('MMPM: Capacity Planning');
    const hasMMPMSheet = mmpmSheet !== null;

    if (!hasMMPMSheet) {
      console.log('MMPM: Capacity Planning sheet not found - using static values');
    }

    // Data rows
    const iterations = [
      'Iteration 1', 'Iteration 2', 'Iteration 3', 'Iteration 4',
      'Iteration 5', 'Iteration 6', 'Total (4 iterations)', 'Total (6 iterations)'
    ];

    iterations.forEach((iteration, index) => {
      const iterationNum = index + 1;

      sheet.getRange(startRow, 1).setValue(iteration);

      if (iterationNum <= 6) {
        // For regular iterations, use safe values if MMPM sheet doesn't exist
        if (hasMMPMSheet) {
          // Try to use MMPM formulas (wrapped in IFERROR)
          sheet.getRange(startRow, 2).setValue(40); // Placeholder baseline capacity
          sheet.getRange(startRow, 3).setValue(30); // Placeholder product load
        } else {
          // Use static values
          sheet.getRange(startRow, 2).setValue(40); // Default capacity
          sheet.getRange(startRow, 3).setValue(30); // Default load
        }

        // Slotted values from actual data
        sheet.getRange(startRow, 4).setValue(slottedData.product[iterationNum] || 0);
        sheet.getRange(startRow, 5).setFormula(`=C${startRow}-D${startRow}`);

        sheet.getRange(startRow, 6).setValue(5); // Default tech load
        sheet.getRange(startRow, 7).setValue(slottedData.tech[iterationNum] || 0);
        sheet.getRange(startRow, 8).setFormula(`=F${startRow}-G${startRow}`);

        sheet.getRange(startRow, 9).setValue(5); // Default quality load
        sheet.getRange(startRow, 10).setValue(slottedData.quality[iterationNum] || 0);
        sheet.getRange(startRow, 11).setFormula(`=I${startRow}-J${startRow}`);
      }

      startRow++;
    });

    return startRow + 2;

  } catch (error) {
    console.error('Error in createSafeIterationSlottingChart:', error);
    return startRow + 2;
  }
}
function checkMMPMSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const mmpmSheet = spreadsheet.getSheetByName('MMPM: Capacity Planning');

  if (mmpmSheet) {
    console.log('MMPM: Capacity Planning sheet EXISTS');

    // Check if it has data in the expected cells
    const testCells = ['B31', 'B42', 'B53', 'B64', 'B75', 'B86'];
    testCells.forEach(cell => {
      const value = mmpmSheet.getRange(cell).getValue();
      console.log(`Cell ${cell}: ${value}`);
    });
  } else {
    console.log('MMPM: Capacity Planning sheet NOT FOUND - This is likely the issue!');
  }
}
function testProgressBarCreation() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const testSheet = spreadsheet.getSheetByName('PI 12 - Time-Keepers Summary');

  if (!testSheet) {
    console.log('Time-Keepers summary sheet not found');
    return;
  }

  try {
    // Try to create a progress bar directly
    console.log('Testing progress bar creation...');
    createProgressBar(testSheet, 8, 4, 75);
    console.log('Progress bar created successfully at row 8, column 4');

    // Check what's in that cell
    const value = testSheet.getRange(8, 4).getValue();
    console.log('Cell value:', value);

  } catch (error) {
    console.error('Error creating progress bar:', error);
  }
}
function regenerateTimeKeepersSkippingProgress() {
  const ui = SpreadsheetApp.getUi();

  try {
    // Override createTeamPlanningProgressGauges to skip the progress bars
    const original = createTeamPlanningProgressGauges;

    createTeamPlanningProgressGauges = function(sheet, startRow, allIssues, epics, stories, scrumTeam, programIncrement) {
      console.log('Using simplified planning progress (no bars)');

      sheet.getRange(startRow, 1).setValue('Planning Progress');
      sheet.getRange(startRow, 1).setFontSize(14).setFontWeight('bold').setBackground('#E1D5E7');
      startRow += 2;

      // Just show the data without progress bars
      const capacityData = getCapacityDataForTeam(sheet.getParent(), scrumTeam);
      const totalStoryPoints = stories.reduce((sum, s) => sum + (s.storyPoints || 0), 0);
      const percentCapacity = capacityData && capacityData.total > 0 ?
        Math.round((totalStoryPoints / capacityData.total) * 100) : 0;

      sheet.getRange(startRow, 1).setValue('Capacity Allocated:');
      sheet.getRange(startRow, 2).setValue(`${percentCapacity}% (${totalStoryPoints} / ${capacityData?.total || 0})`);
      startRow += 2;

      // Continue with iteration slotting
      return createIterationSlottingChart(sheet, startRow, allIssues, scrumTeam, programIncrement);
    };

    // Regenerate
    showProgress('Regenerating Time-Keepers summary...');
    generateSummaryForScrumTeam('12', 'Time-Keepers');
    closeProgress();

    // Restore
    createTeamPlanningProgressGauges = original;

    ui.alert('Time-Keepers summary regenerated successfully!');

  } catch (error) {
    console.error('Error:', error);
    ui.alert('Error', error.toString(), ui.ButtonSet.OK);
  }
}
function regenerateTimeKeepersSummaryFixed() {
  const ui = SpreadsheetApp.getUi();

  try {
    // Delete existing sheet
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const existingSheet = spreadsheet.getSheetByName('PI 12 - Time-Keepers Summary');
    if (existingSheet) {
      spreadsheet.deleteSheet(existingSheet);
    }

    // Regenerate
    showProgress('Regenerating Time-Keepers summary...');
    generateSummaryForScrumTeam('12', 'Time-Keepers');
    closeProgress();

    ui.alert('Success', 'Time-Keepers summary has been regenerated with overallocation handling!', ui.ButtonSet.OK);

  } catch (error) {
    console.error('Error:', error);
    closeProgress();
    ui.alert('Error', error.toString(), ui.ButtonSet.OK);
  }
}
function fixAllExistingSummaryProgressBars() {
  const ui = SpreadsheetApp.getUi();
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = spreadsheet.getSheets();

  let fixedCount = 0;

  sheets.forEach(sheet => {
    const sheetName = sheet.getName();

    // Only process summary sheets
    if (sheetName.includes('Summary')) {
      try {
        // Look for cells with the problematic characters
        const dataRange = sheet.getDataRange();
        const values = dataRange.getValues();

        for (let row = 0; row < values.length; row++) {
          for (let col = 0; col < values[row].length; col++) {
            const cellValue = values[row][col];
            if (cellValue && typeof cellValue === 'string' &&
                (cellValue.includes('█') || cellValue.includes('░'))) {
              // Found a progress bar cell
              const matches = cellValue.match(/█+░*/);
              if (matches) {
                const filled = (cellValue.match(/█/g) || []).length;
                const empty = (cellValue.match(/░/g) || []).length;
                const total = filled + empty;
                const percentage = Math.round((filled / total) * 100);

                // Create new progress bar
                createProgressBar(sheet, row + 1, col + 1, percentage);
                fixedCount++;
              }
            }
          }
        }
      } catch (error) {
        console.error(`Error fixing ${sheetName}:`, error);
      }
    }
  });

  ui.alert('Progress Bars Fixed', `Fixed ${fixedCount} progress bars across all summary sheets`, ui.ButtonSet.OK);
}
function debugTimeKeepersIssue() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const piSheet = spreadsheet.getSheetByName('PI 12');

  if (!piSheet) {
    console.log('No PI 12 sheet found');
    return;
  }

  // Read data
  const dataRange = piSheet.getDataRange();
  const values = dataRange.getValues();
  const headers = values[3];

  const issues = parsePISheetData(values, headers);

  // Filter for Time-Keepers
  const timeKeepersIssues = issues.filter(issue =>
    issue.scrumTeam === 'Time-Keepers' || issue.scrumTeam === 'TIME-KEEPERS'
  );

  console.log('Time-Keepers issues found:', timeKeepersIssues.length);

  const epics = timeKeepersIssues.filter(i => i.issueType === 'Epic');
  const stories = timeKeepersIssues.filter(i => i.issueType === 'Story');

  console.log('Epics:', epics.length);
  console.log('Stories:', stories.length);

  // Check capacity data
  const capacityData = getCapacityDataForTeam(spreadsheet, 'Time-Keepers');
  console.log('Capacity data:', capacityData);

  // Calculate metrics
  const totalStoryPoints = stories.reduce((sum, s) => sum + (s.storyPoints || 0), 0);
  console.log('Total story points:', totalStoryPoints);

  if (capacityData && capacityData.total) {
    const percentage = Math.round((totalStoryPoints / capacityData.total) * 100);
    console.log('Calculated percentage:', percentage);
  } else {
    console.log('No capacity data found');
  }
}
/**
 * Debug function to check capacity calculation for a specific team
 * Run this from the Script Editor or create a menu item for it
 */
function debugCapacityCalculation() {
  const teamName = "TIME-KEEPERS"; // Change this to test different teams
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  console.log("=== DEBUGGING CAPACITY CALCULATION ===");
  console.log(`Team to check: ${teamName}`);

  // Get the capacity sheet
  const capacitySheet = spreadsheet.getSheetByName('Capacity');
  if (!capacitySheet) {
    console.log('ERROR: Capacity sheet not found');
    return;
  }

  // Get all data
  const dataRange = capacitySheet.getDataRange();
  const values = dataRange.getValues();

  console.log(`\nCapacity sheet has ${values.length} rows and ${values[0].length} columns`);

  // Show headers
  if (values.length > 0) {
    console.log("\nHeaders (Row 1):");
    values[0].forEach((header, index) => {
      console.log(`  Column ${String.fromCharCode(65 + index)} (index ${index}): "${header}"`);
    });
  }

  // Normalize team name for matching
  const normalizeTeamName = (name) => {
    return name.toUpperCase().replace(/[\s-]/g, '');
  };

  const normalizedSearchTeam = normalizeTeamName(teamName);
  console.log(`\nNormalized search team: ${normalizedSearchTeam}`);

  // Find the team row
  let teamRow = -1;
  console.log("\nSearching for team...");

  for (let i = 2; i < values.length; i++) {
    const sheetTeamName = values[i][0];
    if (sheetTeamName) {
      const normalizedSheetTeam = normalizeTeamName(sheetTeamName.toString());
      console.log(`  Row ${i + 1}: "${sheetTeamName}" -> normalized: "${normalizedSheetTeam}"`);

      if (normalizedSheetTeam === normalizedSearchTeam) {
        teamRow = i;
        console.log(`  ✓ MATCH FOUND!`);
        break;
      }
    }
  }

  if (teamRow === -1) {
    console.log(`\nERROR: Team ${teamName} not found in capacity sheet`);
    return;
  }

  console.log(`\nFound ${teamName} at row ${teamRow + 1}`);

  // Show all values in the team's row
  console.log("\nAll values in team row:");
  values[teamRow].forEach((value, index) => {
    console.log(`  Column ${String.fromCharCode(65 + index)} (index ${index}): ${value}`);
  });

  // Calculate capacity different ways
  console.log("\n=== CAPACITY CALCULATIONS ===");

  // Method 1: Columns B through F (indices 1-5)
  let totalBtoF = 0;
  console.log("\nMethod 1: Sum of columns B through F");
  for (let col = 1; col <= 5; col++) {
    const value = parseFloat(values[teamRow][col]) || 0;
    totalBtoF += value;
    console.log(`  Column ${String.fromCharCode(65 + col)}: ${value}`);
  }
  console.log(`  Total (B-F): ${totalBtoF}`);

  // Method 2: Columns B through G (indices 1-6) - what the code was doing
  let totalBtoG = 0;
  console.log("\nMethod 2: Sum of columns B through G");
  for (let col = 1; col <= 6; col++) {
    const value = parseFloat(values[teamRow][col]) || 0;
    totalBtoG += value;
    console.log(`  Column ${String.fromCharCode(65 + col)}: ${value}`);
  }
  console.log(`  Total (B-G): ${totalBtoG}`);

  // Now check story points for this team
  console.log("\n=== STORY POINTS CALCULATION ===");

  // Find the PI sheet (assuming PI 12)
  const piNumber = 12; // Change this to your PI number
  const piSheet = spreadsheet.getSheetByName(`PI ${piNumber}`);

  if (!piSheet) {
    console.log(`ERROR: PI ${piNumber} sheet not found`);
    return;
  }

  const piData = piSheet.getDataRange().getValues();
  const piHeaders = piData[3]; // Headers are in row 4 (index 3)

  // Find relevant column indices
  const keyIndex = piHeaders.indexOf('Key');
  const issueTypeIndex = piHeaders.indexOf('Issue Type');
  const scrumTeamIndex = piHeaders.indexOf('Scrum Team');
  const storyPointsIndex = piHeaders.indexOf('Story Points');
  const summaryIndex = piHeaders.indexOf('Summary');

  console.log("\nColumn indices in PI sheet:");
  console.log(`  Key: ${keyIndex}`);
  console.log(`  Issue Type: ${issueTypeIndex}`);
  console.log(`  Scrum Team: ${scrumTeamIndex}`);
  console.log(`  Story Points: ${storyPointsIndex}`);

  // Count story points
  let storyPointsTotal = 0;
  let storyCount = 0;
  let epicCount = 0;
  let otherCount = 0;

  for (let i = 4; i < piData.length; i++) {
    const row = piData[i];
    if (!row[0]) continue; // Skip empty rows

    const issueType = row[issueTypeIndex];
    const scrumTeam = row[scrumTeamIndex] || 'Unassigned';
    const storyPoints = parseFloat(row[storyPointsIndex]) || 0;

    if (scrumTeam === teamName || scrumTeam === "Time-Keepers" || scrumTeam === "TIME-KEEPERS" || scrumTeam === "Time Keepers") {
      if (issueType === 'Epic') {
        epicCount++;
      } else if (issueType === 'Story') {
        storyCount++;
        storyPointsTotal += storyPoints;
      } else {
        otherCount++;
        if (storyPoints > 0) {
          storyPointsTotal += storyPoints;
        }
      }
    }
  }

  console.log(`\nIssues found for ${teamName}:`);
  console.log(`  Epics: ${epicCount}`);
  console.log(`  Stories: ${storyCount}`);
  console.log(`  Other: ${otherCount}`);
  console.log(`  Total Story Points (excluding epics): ${storyPointsTotal}`);

  // Calculate percentages
  console.log("\n=== PERCENTAGE CALCULATIONS ===");
  console.log(`Using capacity B-F (${totalBtoF}):`);
  console.log(`  ${storyPointsTotal} / ${totalBtoF} * 100 = ${Math.round((storyPointsTotal / totalBtoF) * 100)}%`);

  console.log(`\nUsing capacity B-G (${totalBtoG}):`);
  console.log(`  ${storyPointsTotal} / ${totalBtoG} * 100 = ${Math.round((storyPointsTotal / totalBtoG) * 100)}%`);

  // Check if 140% matches any calculation
  const targetPercentage = 140;
  const impliedCapacity = storyPointsTotal / (targetPercentage / 100);
  console.log(`\nTo get ${targetPercentage}%, the capacity would need to be: ${storyPointsTotal} / 1.40 = ${impliedCapacity.toFixed(1)}`);
}
function debugPlanningProgress() {
  const teamName = "TIME-KEEPERS"; // Change this to test different teams
  const piNumber = 12; // Change this to your PI number

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const piSheet = spreadsheet.getSheetByName(`PI ${piNumber}`);

  if (!piSheet) {
    console.log(`ERROR: PI ${piNumber} sheet not found`);
    return;
  }

  console.log("=== DEBUG PLANNING PROGRESS CALCULATION ===");
  console.log(`Team: ${teamName}`);
  console.log(`PI: ${piNumber}`);

  // Read all data from PI sheet
  const dataRange = piSheet.getDataRange();
  const values = dataRange.getValues();
  const headers = values[3]; // Headers in row 4

  // Parse all issues
  const allIssues = [];
  for (let i = 4; i < values.length; i++) {
    const row = values[i];
    if (!row[0]) continue;

    const issue = {};
    headers.forEach((header, index) => {
      const value = row[index];
      switch(header) {
        case 'Key': issue.key = value; break;
        case 'Issue Type': issue.issueType = value; break;
        case 'Summary': issue.summary = value; break;
        case 'Status': issue.status = value; break;
        case 'Story Points': issue.storyPoints = parseFloat(value) || 0; break;
        case 'Feature Points': issue.featurePoints = parseFloat(value) || 0; break;
        case 'Epic Link': issue.epicLink = value; break;
        case 'Parent': issue.parentKey = value; break;
        case 'Scrum Team': issue.scrumTeam = value || 'Unassigned'; break;
        case 'Sprint': issue.sprintName = value; break;
      }
    });
    allIssues.push(issue);
  }

  // Filter for team
  const teamIssues = allIssues.filter(issue =>
    (issue.scrumTeam || 'Unassigned') === teamName
  );

  console.log(`\nTotal issues for ${teamName}: ${teamIssues.length}`);

  // Separate epics and stories (same as the actual function)
  const epics = teamIssues.filter(i => i.issueType === 'Epic');
  const stories = teamIssues.filter(i => i.issueType !== 'Epic');

  console.log(`Epics: ${epics.length}`);
  console.log(`Stories (non-epics): ${stories.length}`);

  // Calculate total story points (same as actual function)
  const totalStoryPoints = stories.reduce((sum, s) => sum + (s.storyPoints || 0), 0);
  console.log(`Total story points from stories: ${totalStoryPoints}`);

  // Check if there are any non-Story issue types contributing points
  const nonStoryWithPoints = stories.filter(s => s.issueType !== 'Story' && s.storyPoints > 0);
  if (nonStoryWithPoints.length > 0) {
    console.log(`\nNon-Story issues with points:`);
    nonStoryWithPoints.forEach(issue => {
      console.log(`  ${issue.key} (${issue.issueType}): ${issue.storyPoints} points`);
    });
  }

  // Get capacity using the same function
  const capacityData = getCapacityDataForTeam(spreadsheet, teamName);

  if (capacityData && capacityData.total > 0) {
    console.log(`\nCapacity from getCapacityDataForTeam: ${capacityData.total}`);

    const percentCapacityAllocated = Math.round((totalStoryPoints / capacityData.total) * 100);
    console.log(`\nCalculation: ${totalStoryPoints} / ${capacityData.total} * 100 = ${percentCapacityAllocated}%`);

    // Check what value would give 140%
    const targetPercent = 140;
    const impliedCapacity = totalStoryPoints / (targetPercent / 100);
    console.log(`\nTo get ${targetPercent}%, capacity would need to be: ${impliedCapacity.toFixed(1)}`);

    // Check if there's a specific calculation difference
    const difference = capacityData.total - impliedCapacity;
    console.log(`Difference: ${difference.toFixed(1)}`);

  } else {
    console.log("\nNo capacity data found!");
  }

  // Let's also check the summary sheet to see what's actually there
  const sheetName = `PI ${piNumber} - ${teamName} Summary`;
  const summarySheet = spreadsheet.getSheetByName(sheetName);

  if (summarySheet) {
    console.log(`\n=== CHECKING SUMMARY SHEET ===`);
    console.log(`Sheet: ${sheetName}`);

    // Look for the Planning Progress section
    const dataRange = summarySheet.getDataRange();
    const values = dataRange.getValues();

    for (let row = 0; row < Math.min(20, values.length); row++) {
      const rowValues = values[row];
      if (rowValues[0] && rowValues[0].toString().includes('% of Planned Capacity Allocated')) {
        console.log(`\nFound '% of Planned Capacity Allocated' at row ${row + 1}`);
        console.log(`Value in column C: ${rowValues[2]}`);

        // Check if it's a formula
        const cell = summarySheet.getRange(row + 1, 3);
        const formula = cell.getFormula();
        if (formula) {
          console.log(`Cell has formula: ${formula}`);
        } else {
          console.log(`Cell has static value: ${cell.getValue()}`);
        }
        break;
      }
    }
  } else {
    console.log(`\nSummary sheet not found: ${sheetName}`);
  }
}

// Helper function needed by the debug
function getCapacityDataForTeam(spreadsheet, teamName) {
  try {
    const capacitySheet = spreadsheet.getSheetByName('Capacity');
    if (!capacitySheet) {
      console.log('Capacity sheet not found');
      return null;
    }

    const dataRange = capacitySheet.getDataRange();
    const values = dataRange.getValues();

    if (values.length < 3) return null;

    // Normalize team name for matching
    const normalizeTeamName = (name) => {
      return name.toUpperCase().replace(/[\s-]/g, '');
    };

    const normalizedSearchTeam = normalizeTeamName(teamName);

    // Find the team row with flexible matching
    let teamRow = -1;
    for (let i = 2; i < values.length; i++) {
      const sheetTeamName = values[i][0];
      if (sheetTeamName) {
        const normalizedSheetTeam = normalizeTeamName(sheetTeamName.toString());
        if (normalizedSheetTeam === normalizedSearchTeam) {
          teamRow = i;
          break;
        }
      }
    }

    if (teamRow === -1) {
      console.log(`Team ${teamName} not found in capacity sheet`);
      return null;
    }

    console.log(`Found team ${teamName} at row ${teamRow + 1} in capacity sheet`);

    // Calculate total capacity from columns B through F
    let total = 0;
    for (let col = 1; col <= 5; col++) {  // Only B through F
      total += parseFloat(values[teamRow][col]) || 0;
    }

    console.log(`Total capacity: ${total}`);
    return { total: total };

  } catch (error) {
    console.error('Error reading capacity data:', error);
    return null;
  }
}

function verifyPIEpics() {
  const ui = SpreadsheetApp.getUi();
  const piNumber = "12";
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const piSheet = spreadsheet.getSheetByName(`PI ${piNumber}`);

  if (!piSheet) {
    ui.alert('No PI 12 sheet found');
    return;
  }

  try {
    showProgress('Checking PI 12 epics against JIRA...');

    // Get all epic keys from the sheet
    const dataRange = piSheet.getDataRange();
    const values = dataRange.getValues();
    const headers = values[3];

    const keyCol = headers.indexOf('Key');
    const issueTypeCol = headers.indexOf('Issue Type');
    const summaryCol = headers.indexOf('Summary');
    const vsCol = headers.indexOf('Value Stream');

    const sheetEpics = [];
    for (let i = 4; i < values.length; i++) {
      if (values[i][issueTypeCol] === 'Epic') {
        sheetEpics.push({
          key: values[i][keyCol],
          summary: values[i][summaryCol],
          valueStream: values[i][vsCol]
        });
      }
    }

    console.log(`Found ${sheetEpics.length} epics in sheet`);

    // Query JIRA for what's actually in PI 12 for MMPM
    const jql = `issuetype = Epic AND cf[10113] = "PI 12" AND cf[10046] = "MMPM" AND status != "Closed"`;
    console.log('JQL Query:', jql);

    const jiraEpics = searchJiraIssues(jql);
    const jiraEpicKeys = new Set(jiraEpics.map(e => e.key));

    console.log(`Found ${jiraEpics.length} MMPM epics in JIRA for PI 12`);

    // Find discrepancies
    const notInJira = sheetEpics.filter(e => !jiraEpicKeys.has(e.key) && e.valueStream === 'MMPM');
    const inJiraButNotSheet = jiraEpics.filter(e => !sheetEpics.find(se => se.key === e.key));

    closeProgress();

    let message = `PI 12 MMPM Epic Verification:\n\n`;
    message += `Sheet has ${sheetEpics.filter(e => e.valueStream === 'MMPM').length} MMPM epics\n`;
    message += `JIRA has ${jiraEpics.length} MMPM epics in PI 12\n\n`;

    if (notInJira.length > 0) {
      message += `\nEpics in sheet but NOT in PI 12 JIRA query:\n`;
      notInJira.forEach(e => {
        message += `- ${e.key}: ${e.summary}\n`;
      });
    }

    if (inJiraButNotSheet.length > 0) {
      message += `\nEpics in JIRA PI 12 but NOT in sheet:\n`;
      inJiraButNotSheet.slice(0, 5).forEach(e => {
        message += `- ${e.key}: ${e.summary}\n`;
      });
    }

    if (notInJira.length === 0 && inJiraButNotSheet.length === 0) {
      message += '\n✅ Sheet and JIRA are in sync!';
    }

    // Now check specific epics that might have moved
    if (notInJira.length > 0) {
      message += '\n\nChecking where these epics went...';

      showProgress('Checking moved epics...');
      const epicKeys = notInJira.map(e => e.key);
      const checkJql = `key in (${epicKeys.join(',')})`;
      const movedEpics = searchJiraIssues(checkJql);

      message += '\n\nMoved epics:\n';
      movedEpics.forEach(epic => {
        message += `- ${epic.key} is now in ${epic.programIncrement || 'NO PI'}\n`;
      });
      closeProgress();
    }

    ui.alert('Verification Results', message, ui.ButtonSet.OK);

  } catch (error) {
    closeProgress();
    ui.alert('Error', 'Verification failed: ' + error.toString(), ui.ButtonSet.OK);
  }
}

/**
 * Force refresh MMPM data for PI 12
 */
function forceRefreshMMPM() {
  const ui = SpreadsheetApp.getUi();

  try {
    // Clear cache first
    clearPICaches("12");

    // Run fresh analysis for MMPM only
    showProgress('Force refreshing MMPM data for PI 12...');
    analyzeSelectedValueStreams("12", ["MMPM"]);

  } catch (error) {
    closeProgress();
    ui.alert('Error', 'Refresh failed: ' + error.toString(), ui.ButtonSet.OK);
  }
}
