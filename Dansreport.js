/**
 * DAN'S REPORT - Clinical Capacity Utilization Snapshot
 * 
 * This is a standalone script that creates and updates a "DAN's Report" tab
 * with a COMBINED capacity utilization snapshot showing ENTIRE PI and CODE FREEZE
 * metrics side-by-side for easy comparison.
 * 
 * User Workflow:
 * 1. User selects which PI to analyze (e.g., PI 13)
 * 2. User chooses whether to refresh PI data first (optional)
 *    - If Yes: Calls analyzeSelectedValueStreams(['EMA Clinical']) to refresh all clinical team data
 *    - If No: Uses existing PI data and generates report immediately
 * 3. Report is generated with combined table showing both perspectives
 * 
 * Data Sources:
 * - Clinical: Capacity Planning tab (for baseline and product capacity)
 *   - Column A: Team names
 *   - Column B: Baseline capacity values
 *   - Column L: Product capacity part 1
 *   - Column M: Product capacity part 2
 *   
 *   ENTIRE PI Section: Reads rows 36-47
 *   CODE FREEZE Section: Reads rows 4-15
 * 
 * - PI Sheet (e.g., "PI 13") for capacity used
 *   - Headers in row 4
 *   - Data starts in row 5
 *   - Counts only Story and Bug types (excludes Epics)
 *   - Filters by Allocation: "Product - Feature" or "Product - Compliance"
 *   - Native work only (assigned to the team, not dependencies)
 * 
 * Table Structure (16 columns):
 * - Column 1: Scrum Team
 * - Column 2: Baseline Capacity (same for both) - LIGHT GREY BACKGROUND
 * - Columns 3-9: ENTIRE PI metrics
 *   - Product Capacity
 *   - Capacity Used (LOE)
 *   - Capacity Remaining
 *   - Planned Load (FP)
 *   - Planned Remaining
 *   - Actual Load (SPE)
 *   - Actual Remaining
 * - Columns 10-16: CODE FREEZE metrics
 *   - Product Capacity
 *   - Capacity Used (LOE)
 *   - Capacity Remaining
 *   - Planned Load (FP)
 *   - Planned Remaining
 *   - Actual Load (SPE)
 *   - Actual Remaining
 * 
 * Visual Grouping:
 * - Thicker borders (SOLID_MEDIUM) around paired variance columns:
 *   - Capacity Used (LOE) + Capacity Remaining
 *   - Planned Load (FP) + Planned Remaining
 *   - Actual Load (SPE) + Actual Remaining
 * - This visual grouping appears in both ENTIRE PI and CODE FREEZE sections
 * - CODE FREEZE columns with different data source have LIGHT BLUE BACKGROUND:
 *   - Product Capacity (column 10)
 *   - Capacity Remaining (column 12)
 *   - Planned Remaining (column 14)
 *   - Actual Remaining (column 16)
 * - Columns without blue (Capacity Used, Planned Load, Actual Load) are the same in both sections
 * 
 * Special Rows:
 * - SUBTOTAL (Excluding EYEFINITY): Appears before EYEFINITY row, sums all teams except EYEFINITY
 * - EYEFINITY: Shows EYEFINITY team data
 * - TOTAL (Including EYEFINITY): Final row = SUBTOTAL + EYEFINITY only (sums these 2 rows for cleaner formulas)
 * 
 * Calculations:
 * - All numbers rounded UP to next whole number using Math.ceil()
 * - Baseline Capacity: From Column B (rounded up)
 * - Product Capacity: Column L + Column M (each rounded up)
 * - Capacity Used: Sum of Story Points for Stories and Bugs with Product allocations (each rounded up)
 * - Capacity Remaining: Product Capacity - Capacity Used
 *   - RED if negative (over capacity)
 *   - GREEN if positive (under capacity)
 * 
 * Load Metrics (calculated separately for ENTIRE PI and CODE FREEZE):
 * - Planned Load (FP): At Epic level, sum Feature Points Ã— 10 where:
 *   - Issue Type = "Epic"
 *   - Allocation = "Product - Feature" OR "Product - Compliance"
 *   - Scrum Team = team in context
 * - Planned Remaining: Product Capacity - Planned Load (FP)
 *   - Uses ENTIRE PI Product Capacity for ENTIRE PI section
 *   - Uses CODE FREEZE Baseline Capacity for CODE FREEZE section
 * - Actual Load (SPE): At Epic level, sum Story Point Estimate where:
 *   - Issue Type = "Epic"
 *   - Allocation = "Product - Feature" OR "Product - Compliance"
 *   - Scrum Team = team in context
 * - Actual Remaining: Product Capacity - Actual Load (SPE)
 *   - Uses ENTIRE PI Product Capacity for ENTIRE PI section
 *   - Uses CODE FREEZE Baseline Capacity for CODE FREEZE section
 */

// ===== CONFIGURATION =====

const DANS_REPORT_CONFIG = {
  reportSheetName: "DAN's Report",
  capacitySheetName: "Clinical: Capacity Planning",
  
  // JIRA base URL for creating hyperlinks
  jiraBaseUrl: 'https://modmedrnd.atlassian.net',
  
  // Clinical scrum teams to include (in display order, UPPERCASE for matching)
  // Used for case-insensitive matching when reading PI data
  clinicalTeams: [
    'ORDERNAUTS',
    'EMBRYONICS',
    'ALCHEMIST',
    'VESTIES',
    'SPICE RUNNERS',
    'MANDALORE',
    'PATIENCE',
    'AVENGERS',
    'EXPLORERS',
    'ARTIFICIALLY INTELLIGENT',
    'PAIN KILLERS',
    'EYEFINITY'
  ],
  
  // Capacity sheet locations (1-indexed)
  capacityRows: {
    entirePI: {
      start: 36,
      end: 47
    },
    codeFreeze: {
      start: 4,
      end: 15
    }
  },
  capacityColumns: {
    teamName: 'A',        // Column A - Team names
    baseline: 'B',        // Column B - Baseline capacity
    productL: 'L',        // Column L - Product capacity part 1
    productM: 'M'         // Column M - Product capacity part 2
  }
};

// Role normalization map - maps various role name variants to canonical keys
const ROLE_NORMALIZATION_DANS = {
  'QA': 'QA',
  'AQA': 'QA',
  'W-DEV': 'W-DEV',
  'WDEV': 'W-DEV',
  'W DEV': 'W-DEV',
  'M-DEV': 'M-DEV',
  'MDEV': 'M-DEV',
  'M DEV': 'M-DEV',
  'MOBILE': 'M-DEV',
  'M-ANDROID': 'M-ANDROID',
  'M-IOS': 'M-IOS',
  'BE': 'BE',
  'FE': 'FE',
  'DEVOPS': 'DEVOPS',
  'UX': 'UX'
};

// ===== MAIN REPORT GENERATION =====

/**
 * Main entry point to create or update DAN's Report
 * Can be called from menu or directly
 */
function generateDansReport() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    // Prompt for PI number
    const piResponse = ui.prompt(
      'DAN\'s Report - PI Selection',
      'Enter PI number (e.g., 11, 12, 13):',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (piResponse.getSelectedButton() !== ui.Button.OK) {
      return; // User cancelled
    }
    
    const piNumber = piResponse.getResponseText().trim();
    if (!piNumber || !/^\d+$/.test(piNumber)) {
      ui.alert('Invalid PI format. Please use a number like "11", "12", or "13"');
      return;
    }
    
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const programIncrement = `PI ${piNumber}`;
    
    // Ask if user wants to refresh PI data first
    const refreshResponse = ui.alert(
      'Refresh PI Data?',
      `Do you want to refresh the data for ${programIncrement} before generating the report?\n\n` +
      `This will re-analyze ${programIncrement} for the EMA Clinical value stream (includes all clinical scrum teams).\n\n` +
      `Choose "Yes" to refresh data (takes longer) or "No" to use existing data.`,
      ui.ButtonSet.YES_NO
    );
    
    if (refreshResponse === ui.Button.YES) {
      try {
        // Call the analysis function directly with EMA Clinical value stream
        // This will fetch all clinical team data and show its own progress dialogs
        console.log(`Calling analyzeSelectedValueStreams for ${programIncrement} with EMA Clinical value stream`);
        
        // Check if the function exists
        if (typeof analyzeSelectedValueStreams === 'function') {
          // Call with EMA Clinical value stream (which includes all clinical teams)
          // This function will handle its own UI alerts and progress indicators
          analyzeSelectedValueStreams(piNumber, ['EMA Clinical']);
          
          ui.alert(
            'Data Refresh Complete',
            `${programIncrement} data has been refreshed for EMA Clinical value stream.\n\n` +
            `Now generating DAN's Report...`,
            ui.ButtonSet.OK
          );
        } else {
          ui.alert(
            'Function Not Found',
            `The analyzeSelectedValueStreams function is not available.\n\n` +
            `Proceeding with existing data...`,
            ui.ButtonSet.OK
          );
        }
      } catch (error) {
        console.error('Error refreshing PI data:', error);
        ui.alert(
          'Refresh Error',
          `Error refreshing data: ${error.toString()}\n\n` +
          `Proceeding with existing data...`,
          ui.ButtonSet.OK
        );
      }
    }
    
    // Check if PI data sheet exists
    const piSheetName = programIncrement;
    const piSheet = spreadsheet.getSheetByName(piSheetName);
    
    if (!piSheet) {
      ui.alert('No PI Data Found', 
               `Sheet "${piSheetName}" not found. Please run the PI analysis first.`,
               ui.ButtonSet.OK);
      return;
    }
    
    // Create or get the report sheet
    let reportSheet = spreadsheet.getSheetByName(DANS_REPORT_CONFIG.reportSheetName);
    if (!reportSheet) {
      reportSheet = spreadsheet.insertSheet(DANS_REPORT_CONFIG.reportSheetName);
    }
    
    // Clear existing content
    reportSheet.clear();
    
    // Get capacity data for ENTIRE PI (rows 36-47) from Clinical: Capacity Planning
    console.log('Reading Entire PI capacity data (rows 36-47)...');
    const capacityDataEntirePI = getClinicalCapacityData(spreadsheet, 36, 47);
    
    // Get capacity data for CODE FREEZE (rows 4-15) from Clinical: Capacity Planning
    console.log('Reading Code Freeze capacity data (rows 4-15)...');
    const capacityDataCodeFreeze = getClinicalCapacityData(spreadsheet, 4, 15);
    
    // Get PI data for capacity used calculations
    const piData = getPIDataForCapacityUsed(spreadsheet, piSheet);
    
    // Generate the report with both tables and blank fix version table
    createCapacityUtilizationReport(reportSheet, capacityDataEntirePI, capacityDataCodeFreeze, piData, programIncrement, piSheet, piNumber);
    
    // Format the sheet
    formatDansReportSheet(reportSheet);
    
    ui.alert('Success!', 
             `DAN's Report has been generated successfully for ${programIncrement}.`, 
             ui.ButtonSet.OK);
    
  } catch (error) {
    console.error('Error generating DAN\'s Report:', error);
    ui.alert('Error', 
             `Failed to generate DAN's Report:\n${error.toString()}`, 
             ui.ButtonSet.OK);
  }
}

// ===== DATA COLLECTION FUNCTIONS =====

/**
 * Get capacity data from the Clinical: Capacity Planning sheet
 * Returns object with baseline and product capacity for each team
 * @param {Spreadsheet} spreadsheet - The spreadsheet object
 * @param {number} startRow - Starting row number (e.g., 36 for Entire PI, 4 for Code Freeze)
 * @param {number} endRow - Ending row number (e.g., 47 for Entire PI, 15 for Code Freeze)
 */
function getClinicalCapacityData(spreadsheet, startRow, endRow) {
  const capacitySheet = spreadsheet.getSheetByName(DANS_REPORT_CONFIG.capacitySheetName);
  
  if (!capacitySheet) {
    throw new Error(`Capacity sheet "${DANS_REPORT_CONFIG.capacitySheetName}" not found`);
  }
  
  const config = DANS_REPORT_CONFIG;
  const capacityData = {};
  
  console.log(`Reading capacity data from rows ${startRow} to ${endRow}`);
  
  // Get team names from column A
  const teamRange = capacitySheet.getRange(`${config.capacityColumns.teamName}${startRow}:${config.capacityColumns.teamName}${endRow}`);
  const teamNames = teamRange.getValues().map(row => row[0]);
  
  // Get baseline capacity from column B
  const baselineRange = capacitySheet.getRange(`${config.capacityColumns.baseline}${startRow}:${config.capacityColumns.baseline}${endRow}`);
  const baselineValues = baselineRange.getValues().map(row => row[0]);
  
  // Get product capacity columns L and M
  const productLRange = capacitySheet.getRange(`${config.capacityColumns.productL}${startRow}:${config.capacityColumns.productL}${endRow}`);
  const productLValues = productLRange.getValues().map(row => row[0]);
  
  const productMRange = capacitySheet.getRange(`${config.capacityColumns.productM}${startRow}:${config.capacityColumns.productM}${endRow}`);
  const productMValues = productMRange.getValues().map(row => row[0]);
  
  // Process each row
  for (let i = 0; i < teamNames.length; i++) {
    const teamName = teamNames[i];
    
    if (!teamName || teamName.toString().trim() === '') {
      continue;
    }
    
    // Convert to uppercase for case-insensitive matching
    const cleanTeamName = teamName.toString().trim().toUpperCase();
    
    // Parse numeric values and round up to next whole number
    const baseline = Math.ceil(parseFloat(baselineValues[i]) || 0);
    const productL = Math.ceil(parseFloat(productLValues[i]) || 0);
    const productM = Math.ceil(parseFloat(productMValues[i]) || 0);
    const productCapacity = productL + productM; // Already rounded since L and M are rounded
    
    capacityData[cleanTeamName] = {
      baseline: baseline,
      productCapacity: productCapacity,
      productL: productL,
      productM: productM
    };
    
    console.log(`  Team: ${cleanTeamName}, Baseline: ${baseline}, Product: ${productCapacity} (L:${productL} + M:${productM})`);
  }
  
  return capacityData;
}

/**
 * Get PI data to calculate capacity used (story points) for each team
 * 
 * LOGIC FOR CAPACITY USED:
 * 1. Reads from the PI sheet (e.g., "PI 13")
 * 2. Starts from row 5 (data rows after headers in row 4)
 * 3. For each row:
 *    - Checks if Scrum Team matches a clinical team (case-insensitive)
 *    - Checks if Issue Type is "Story" or "Bug" (excludes Epics)
 *    - Checks if Allocation is "Product - Feature" or "Product - Compliance"
 *    - If ALL conditions met, adds Story Points to the team's total
 * 4. Rounds up each story point value to next whole number
 * 5. Returns sum of all Story Points for each clinical team
 * 
 * NOTE: This counts NATIVE WORK only (stories/bugs assigned to the team)
 *       with Product - Feature or Product - Compliance allocation
 *       Does NOT include dependencies, other allocations, or work assigned to other teams
 * 
 * @param {Spreadsheet} spreadsheet - The spreadsheet object
 * @param {Sheet} piSheet - The PI data sheet to read from
 * @return {Object} Object with team names (uppercase) as keys and total story points as values
 */
function getPIDataForCapacityUsed(spreadsheet, piSheet) {
  if (!piSheet) {
    console.warn('PI sheet not provided - capacity used will be 0');
    return {};
  }
  
  console.log(`Reading PI data from sheet: ${piSheet.getName()}`);
  
  const dataRange = piSheet.getDataRange();
  const values = dataRange.getValues();
  
  if (values.length < 4) {
    console.warn('PI sheet has insufficient data (need at least 4 rows)');
    return {};
  }
  
  // Headers are in row 4 (index 3)
  const headers = values[3];
  console.log('Headers found:', headers.join(', '));
  
  // Find column indices
  const scrumTeamCol = headers.indexOf('Scrum Team');
  const storyPointsCol = headers.indexOf('Story Points');
  const issueTypeCol = headers.indexOf('Issue Type');
  const allocationCol = headers.indexOf('Allocation');
  
  if (scrumTeamCol === -1) {
    console.error('Scrum Team column not found in PI sheet');
    return {};
  }
  
  if (storyPointsCol === -1) {
    console.error('Story Points column not found in PI sheet');
    return {};
  }
  
  console.log(`Column indices - Scrum Team: ${scrumTeamCol}, Story Points: ${storyPointsCol}, Issue Type: ${issueTypeCol}, Allocation: ${allocationCol}`);
  
  // Calculate story points per team (starting from row 5, index 4)
  // Use uppercase keys for case-insensitive matching
  const teamCapacityUsed = {};
  
  for (let i = 4; i < values.length; i++) {
    const row = values[i];
    const scrumTeam = row[scrumTeamCol];
    const storyPoints = parseFloat(row[storyPointsCol]) || 0;
    const issueType = issueTypeCol !== -1 ? row[issueTypeCol] : '';
    const allocation = allocationCol !== -1 ? row[allocationCol] : '';
    
    // Skip empty rows
    if (!scrumTeam || scrumTeam.toString().trim() === '') {
      continue;
    }
    
    // Convert to uppercase for case-insensitive matching
    const cleanTeamName = scrumTeam.toString().trim().toUpperCase();
    
    // Only count Story and Bug types with Product - Feature or Product - Compliance allocation
    if ((issueType === 'Story' || issueType === 'Bug') && 
        (allocation === 'Product - Feature' || allocation === 'Product - Compliance')) {
      if (!teamCapacityUsed[cleanTeamName]) {
        teamCapacityUsed[cleanTeamName] = 0;
      }
      // Round up the story points before adding
      const roundedPoints = Math.ceil(storyPoints);
      teamCapacityUsed[cleanTeamName] += roundedPoints;
      
      // Log first few entries for debugging
      if (Object.keys(teamCapacityUsed).length <= 5) {
        console.log(`  Adding ${roundedPoints} points (from ${storyPoints}) to ${cleanTeamName} (${issueType}, ${allocation})`);
      }
    }
  }
  
  console.log('\n=== Capacity Used Summary ===');
  Object.keys(teamCapacityUsed).sort().forEach(team => {
    console.log(`  ${team}: ${teamCapacityUsed[team]} points`);
  });
  console.log('=== End Capacity Used ===\n');
  
  return teamCapacityUsed;
}

/**
 * Get load data for capacity tables (Planned Load and Actual Load)
 * This calculates load at the EPIC level for both Entire PI and Code Freeze tables
 * 
 * FILTERS:
 * - Issue Type = "Epic"
 * - Allocation = "Product - Feature" OR "Product - Compliance"
 * - NO FIX VERSION FILTER - sums all epics regardless of fix version
 * 
 * CALCULATIONS:
 * - Planned Load (FP): Sum of Feature Points Ã— 10
 * - Actual Load (SPE): Sum of Story Point Estimate
 * 
 * @param {Spreadsheet} spreadsheet - The spreadsheet object
 * @param {Sheet} piSheet - The PI data sheet to read from
 * @param {string} piNumber - The PI number (e.g., "13")
 * @return {Object} Object with two sub-objects: anticipatedLoad and actualLoad, keyed by team name (uppercase)
 */
function getLoadData(spreadsheet, piSheet, piNumber) {
  if (!piSheet) {
    console.warn('PI sheet not provided - load data will be 0');
    return { anticipatedLoad: {}, actualLoad: {} };
  }
  
  console.log(`\n=== Calculating Load Data for PI ${piNumber} ===`);
  console.log('Filtering: Issue Type = Epic, Allocation = Product - Feature OR Product - Compliance');
  console.log('Planned Load: Sum(Feature Points) Ã— 10');
  console.log('Actual Load: Sum(Story Point Estimate)');
  
  const dataRange = piSheet.getDataRange();
  const values = dataRange.getValues();
  
  if (values.length < 4) {
    console.warn('PI sheet has insufficient data (need at least 4 rows)');
    return { anticipatedLoad: {}, actualLoad: {} };
  }
  
  // Headers are in row 4 (index 3)
  const headers = values[3];
  
  // Find column indices
  const scrumTeamCol = headers.indexOf('Scrum Team');
  const issueTypeCol = headers.indexOf('Issue Type');
  const allocationCol = headers.indexOf('Allocation');
  const featurePointsCol = headers.indexOf('Feature Points');
  const storyPointEstimateCol = headers.indexOf('Story Point Estimate');
  
  if (scrumTeamCol === -1) {
    console.error('Scrum Team column not found in PI sheet');
    return { anticipatedLoad: {}, actualLoad: {} };
  }
  
  if (issueTypeCol === -1) {
    console.error('Issue Type column not found in PI sheet');
    return { anticipatedLoad: {}, actualLoad: {} };
  }
  
  if (allocationCol === -1) {
    console.error('Allocation column not found in PI sheet');
    return { anticipatedLoad: {}, actualLoad: {} };
  }
  
  console.log(`Column indices - Scrum Team: ${scrumTeamCol}, Issue Type: ${issueTypeCol}, Allocation: ${allocationCol}, Feature Points: ${featurePointsCol}, Story Point Estimate: ${storyPointEstimateCol}`);
  
  const anticipatedLoad = {}; // Feature Points Ã— 10
  const actualLoad = {}; // Story Point Estimate
  
  let rowsProcessed = 0;
  let rowsFiltered = 0;
  let fpCount = 0;
  let speCount = 0;
  
  // Process each row starting from row 5 (index 4)
  for (let i = 4; i < values.length; i++) {
    const row = values[i];
    const scrumTeam = row[scrumTeamCol];
    const issueType = row[issueTypeCol];
    const allocation = row[allocationCol];
    
    // Skip empty rows
    if (!scrumTeam || scrumTeam.toString().trim() === '') {
      continue;
    }
    
    rowsProcessed++;
    
    // FILTER 1: Must be an Epic
    if (issueType !== 'Epic') {
      continue;
    }
    
    // FILTER 2: Must be Product - Feature OR Product - Compliance
    if (allocation !== 'Product - Feature' && allocation !== 'Product - Compliance') {
      continue;
    }
    
    rowsFiltered++;
    
    // Get the field values
    const featurePoints = featurePointsCol !== -1 ? parseFloat(row[featurePointsCol]) || 0 : 0;
    const storyPointEstimate = storyPointEstimateCol !== -1 ? parseFloat(row[storyPointEstimateCol]) || 0 : 0;
    
    // Convert to uppercase for case-insensitive matching
    const cleanTeamName = scrumTeam.toString().trim().toUpperCase();
    
    // Initialize if needed
    if (!anticipatedLoad[cleanTeamName]) {
      anticipatedLoad[cleanTeamName] = 0;
    }
    if (!actualLoad[cleanTeamName]) {
      actualLoad[cleanTeamName] = 0;
    }
    
    // Add Feature Points Ã— 10 if > 0
    if (featurePoints > 0) {
      const roundedFP = Math.ceil(featurePoints);
      anticipatedLoad[cleanTeamName] += roundedFP * 10; // Multiply by 10
      fpCount++;
    }
    
    // Add Story Point Estimate if > 0
    if (storyPointEstimate > 0) {
      actualLoad[cleanTeamName] += Math.ceil(storyPointEstimate);
      speCount++;
    }
  }
  
  console.log(`\nProcessed ${rowsProcessed} total rows`);
  console.log(`Matched filters: ${rowsFiltered} Epic rows with Product - Feature/Compliance allocation`);
  console.log(`Found ${fpCount} epics with Feature Points > 0`);
  console.log(`Found ${speCount} epics with Story Point Estimate > 0`);
  
  console.log('\n=== Planned Load Summary (FP Ã— 10) ===');
  Object.keys(anticipatedLoad).sort().forEach(team => {
    if (anticipatedLoad[team] > 0) {
      console.log(`  ${team}: ${anticipatedLoad[team]}`);
    }
  });
  
  console.log('\n=== Actual Load Summary (SPE) ===');
  Object.keys(actualLoad).sort().forEach(team => {
    if (actualLoad[team] > 0) {
      console.log(`  ${team}: ${actualLoad[team]}`);
    }
  });
  console.log('=== End Load Data ===\n');
  
  return { anticipatedLoad, actualLoad };
}

/**
 * Create combined capacity utilization table with Entire PI and Code Freeze side-by-side
 * @param {Sheet} sheet - The report sheet
 * @param {number} currentRow - Starting row for this table
 * @param {Object} capacityDataEntirePI - Capacity data for Entire PI (rows 36-47)
 * @param {Object} capacityDataCodeFreeze - Capacity data for Code Freeze (rows 4-15)
 * @param {Object} piData - PI data for capacity used
 * @param {Object} loadData - Load data (anticipatedLoad and actualLoad)
 * @return {number} Next available row after the table
 */
function createCombinedCapacityTable(sheet, currentRow, capacityDataEntirePI, capacityDataCodeFreeze, piData, loadData) {
  const numColumns = 16; // Combined table has 16 columns (1 + 1 + 7 + 7)
  
  // Add section title
  sheet.getRange(currentRow, 1).setValue('CLINICAL CAPACITY UTILIZATION SNAPSHOT - ENTIRE PI vs CODE FREEZE');
  sheet.getRange(currentRow, 1)
    .setFontSize(14)
    .setFontWeight('bold')
    .setFontFamily('Comfortaa')
    .setBackground('#E1D5E7')
    .setFontColor('black');
  sheet.getRange(currentRow, 1, 1, numColumns).merge();
  currentRow++;
  
  // Add category headers (row 1 of headers)
  sheet.getRange(currentRow, 1).setValue('');
  sheet.getRange(currentRow, 2).setValue('');
  sheet.getRange(currentRow, 3, 1, 7).merge();
  sheet.getRange(currentRow, 3).setValue('ENTIRE PI');
  sheet.getRange(currentRow, 10, 1, 7).merge();
  sheet.getRange(currentRow, 10).setValue('CODE FREEZE');
  
  // Format category headers
  sheet.getRange(currentRow, 1, 1, numColumns)
    .setFontWeight('bold')
    .setBackground('#D5A6E0')
    .setFontColor('black')
    .setFontSize(10)
    .setFontFamily('Comfortaa')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  
  currentRow++;
  
  // Add detailed column headers (row 2 of headers)
  const headers = [
    'Scrum Team',
    'Sprint Capacity',
    'Baseline Capacity',
    'Capacity Used (LOE)',
    'Capacity Remaining',
    'Planned Load (FP)',
    'Planned Remaining',
    'Actual Load (SPE)',
    'Actual Remaining',
    'Baseline Capacity',
    'Capacity Used (LOE)',
    'Capacity Remaining',
    'Planned Load (FP)',
    'Planned Remaining',
    'Actual Load (SPE)',
    'Actual Remaining'
  ];
  
  sheet.getRange(currentRow, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(currentRow, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#9b7bb8')
    .setFontColor('white')
    .setFontSize(9)
    .setFontFamily('Comfortaa')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true);
  
  // Add notes to specific column headers
  // Notes are added only to the ENTIRE PI section to avoid redundancy
  // Product Capacity columns - explain where the data comes from
  sheet.getRange(currentRow, 3).setNote('From Clinical: Capacity Planning sheet\nRows 36-47 (ENTIRE PI) or 4-15 (CODE FREEZE)');
  sheet.getRange(currentRow, 10).setNote('From Clinical: Capacity Planning sheet\nRows 36-47 (ENTIRE PI) or 4-15 (CODE FREEZE)\n\nLIGHT BLUE BACKGROUND = Different data source\n(CODE FREEZE capacity from rows 4-15)');
  
  // Capacity Used (LOE) - explain what this measures
  sheet.getRange(currentRow, 4).setNote('LOE = Level of Effort\nSum of story points for Stories and Bugs with Product - Feature or Product - Compliance allocation');
  
  // Capacity Remaining - explain conditional formatting
  sheet.getRange(currentRow, 5).setNote('Baseline Capacity minus Capacity Used\nRED = Over capacity (negative)\nGREEN = Under capacity (positive)');
  sheet.getRange(currentRow, 12).setNote('Baseline Capacity minus Capacity Used\nPINK-BLUE = Over capacity (negative)\nGREEN-BLUE = Under capacity (positive)\n\nLIGHT BLUE BACKGROUND = Uses CODE FREEZE capacity');
  
  // Planned Load (FP) - explain the calculation
  sheet.getRange(currentRow, 6).setNote('At Epic level: sum of (Feature Points Ã— 10)\nFiltered by Product - Feature or Product - Compliance allocation');
  
  // Actual Load (SPE) - explain the calculation
  sheet.getRange(currentRow, 8).setNote('At Epic level: sum of Story Point Estimates\nFiltered by Product - Feature or Product - Compliance allocation');
  
  // Add notes to CODE FREEZE remaining columns explaining the blue tint
  sheet.getRange(currentRow, 14).setNote('CODE FREEZE Baseline Capacity minus Planned Load\nPINK-BLUE = Over capacity (negative)\nGREEN-BLUE = Under capacity (positive)\n\nLIGHT BLUE BACKGROUND = Uses CODE FREEZE capacity');
  sheet.getRange(currentRow, 16).setNote('CODE FREEZE Baseline Capacity minus Actual Load\nPINK-BLUE = Over capacity (negative)\nGREEN-BLUE = Under capacity (positive)\n\nLIGHT BLUE BACKGROUND = Uses CODE FREEZE capacity');
  
  const headerRow = currentRow;
  currentRow++;
  
  // Add data rows for each clinical team
  const dataStartRow = currentRow;
  let eyefinityRowNumber = -1;
  let subtotalRowNumber = -1; // Track the subtotal row number
  
  for (let i = 0; i < DANS_REPORT_CONFIG.clinicalTeams.length; i++) {
    const teamName = DANS_REPORT_CONFIG.clinicalTeams[i];
    
    // Check if this is EYEFINITY - if so, add subtotal row BEFORE it
    if (teamName === 'EYEFINITY' && i > 0) {
      // Add SUBTOTAL row (excluding EYEFINITY)
      const subtotalRow = currentRow;
      subtotalRowNumber = subtotalRow; // Store for later use in TOTAL row
      sheet.getRange(subtotalRow, 1).setValue('SUBTOTAL (Excluding EYEFINITY)');
      
      // Calculate subtotal formulas (from dataStartRow to row before current)
      const subtotalEndRow = currentRow - 1;
      sheet.getRange(subtotalRow, 2).setFormula(`=SUM(B${dataStartRow}:B${subtotalEndRow})`);
      sheet.getRange(subtotalRow, 3).setFormula(`=SUM(C${dataStartRow}:C${subtotalEndRow})`);
      sheet.getRange(subtotalRow, 4).setFormula(`=SUM(D${dataStartRow}:D${subtotalEndRow})`);
      sheet.getRange(subtotalRow, 5).setFormula(`=SUM(E${dataStartRow}:E${subtotalEndRow})`);
      sheet.getRange(subtotalRow, 6).setFormula(`=SUM(F${dataStartRow}:F${subtotalEndRow})`);
      sheet.getRange(subtotalRow, 7).setFormula(`=SUM(G${dataStartRow}:G${subtotalEndRow})`);
      sheet.getRange(subtotalRow, 8).setFormula(`=SUM(H${dataStartRow}:H${subtotalEndRow})`);
      sheet.getRange(subtotalRow, 9).setFormula(`=SUM(I${dataStartRow}:I${subtotalEndRow})`);
      sheet.getRange(subtotalRow, 10).setFormula(`=SUM(J${dataStartRow}:J${subtotalEndRow})`);
      sheet.getRange(subtotalRow, 11).setFormula(`=SUM(K${dataStartRow}:K${subtotalEndRow})`);
      sheet.getRange(subtotalRow, 12).setFormula(`=SUM(L${dataStartRow}:L${subtotalEndRow})`);
      sheet.getRange(subtotalRow, 13).setFormula(`=SUM(M${dataStartRow}:M${subtotalEndRow})`);
      sheet.getRange(subtotalRow, 14).setFormula(`=SUM(N${dataStartRow}:N${subtotalEndRow})`);
      sheet.getRange(subtotalRow, 15).setFormula(`=SUM(O${dataStartRow}:O${subtotalEndRow})`);
      sheet.getRange(subtotalRow, 16).setFormula(`=SUM(P${dataStartRow}:P${subtotalEndRow})`);
      
      // Format subtotal row
      sheet.getRange(subtotalRow, 1, 1, numColumns)
        .setFontWeight('bold')
        .setFontFamily('Comfortaa')
        .setBackground('#E8DAEF')
        .setFontStyle('italic');
      
      sheet.getRange(subtotalRow, 2, 1, 15)
        .setHorizontalAlignment('center')
        .setNumberFormat('#,##0');
      
      // Keep Baseline Capacity light grey even in subtotal row
      sheet.getRange(subtotalRow, 2).setBackground('#e0e0e0'); // Slightly darker grey for subtotal
      
      // Apply light blue tint to CODE FREEZE columns with different data
      sheet.getRange(subtotalRow, 10).setBackground('#d5e5f5'); // Product Capacity - darker blue for subtotal
      sheet.getRange(subtotalRow, 12).setBackground('#d5e5f5'); // Capacity Remaining
      sheet.getRange(subtotalRow, 14).setBackground('#d5e5f5'); // Planned Remaining
      sheet.getRange(subtotalRow, 16).setBackground('#d5e5f5'); // Actual Remaining
      
      currentRow++;
    }
    
    // Track EYEFINITY row
    if (teamName === 'EYEFINITY') {
      eyefinityRowNumber = currentRow;
    }
    
    // Get capacity data from both datasets
    const capacityEntirePI = capacityDataEntirePI[teamName] || { baseline: 0, productCapacity: 0 };
    const capacityCodeFreeze = capacityDataCodeFreeze[teamName] || { baseline: 0, productCapacity: 0 };
    const capacityUsed = piData[teamName] || 0;
    
    // Calculate remaining values for capacity
    const remainingEntirePI = capacityEntirePI.productCapacity - capacityUsed;
    const remainingCodeFreeze = capacityCodeFreeze.productCapacity - capacityUsed;
    
    // Get load data for this team
    const anticipatedLoad = loadData && loadData.anticipatedLoad ? (loadData.anticipatedLoad[teamName] || 0) : 0;
    const actualLoad = loadData && loadData.actualLoad ? (loadData.actualLoad[teamName] || 0) : 0;
    
    // Calculate load remaining for ENTIRE PI
    const anticipatedRemainingEntirePI = capacityEntirePI.productCapacity - anticipatedLoad;
    const actualRemainingEntirePI = capacityEntirePI.productCapacity - actualLoad;
    
    // Calculate load remaining for CODE FREEZE
    const anticipatedRemainingCodeFreeze = capacityCodeFreeze.productCapacity - anticipatedLoad;
    const actualRemainingCodeFreeze = capacityCodeFreeze.productCapacity - actualLoad;
    
    // Write data to all 14 columns
    sheet.getRange(currentRow, 1).setValue(teamName);
    sheet.getRange(currentRow, 2).setValue(capacityEntirePI.baseline);
    
    // ENTIRE PI section (columns 3-9)
    sheet.getRange(currentRow, 3).setValue(capacityEntirePI.productCapacity);
    sheet.getRange(currentRow, 4).setValue(capacityUsed);
    sheet.getRange(currentRow, 5).setValue(remainingEntirePI);
    sheet.getRange(currentRow, 6).setValue(anticipatedLoad);
    sheet.getRange(currentRow, 7).setValue(anticipatedRemainingEntirePI);
    sheet.getRange(currentRow, 8).setValue(actualLoad);
    sheet.getRange(currentRow, 9).setValue(actualRemainingEntirePI);
    
    // CODE FREEZE section (columns 10-16, but we only go to 14 since that's what we have)
    sheet.getRange(currentRow, 10).setValue(capacityCodeFreeze.productCapacity);
    sheet.getRange(currentRow, 11).setValue(capacityUsed);
    sheet.getRange(currentRow, 12).setValue(remainingCodeFreeze);
    sheet.getRange(currentRow, 13).setValue(anticipatedLoad);
    sheet.getRange(currentRow, 14).setValue(anticipatedRemainingCodeFreeze);
    sheet.getRange(currentRow, 15).setValue(actualLoad);
    sheet.getRange(currentRow, 16).setValue(actualRemainingCodeFreeze);
    
    // FORMAT all columns
    sheet.getRange(currentRow, 1, 1, 16)
      .setFontSize(9)
      .setFontFamily('Comfortaa')
      .setVerticalAlignment('middle');
    
    // CENTER JUSTIFY all number columns (2-16)
    sheet.getRange(currentRow, 2, 1, 15)
      .setHorizontalAlignment('center')
      .setNumberFormat('#,##0');
    
    // LIGHT GREY background for Baseline Capacity (column 2)
    sheet.getRange(currentRow, 2).setBackground('#f5f5f5');
    
    // LIGHT BLUE background for CODE FREEZE columns with DIFFERENT data source
    // These columns use CODE FREEZE capacity (rows 4-15) instead of ENTIRE PI capacity
    sheet.getRange(currentRow, 10).setBackground('#e6f2ff'); // Product Capacity
    sheet.getRange(currentRow, 12).setBackground('#e6f2ff'); // Capacity Remaining (calculated with CF capacity)
    sheet.getRange(currentRow, 14).setBackground('#e6f2ff'); // Planned Remaining (calculated with CF capacity)
    sheet.getRange(currentRow, 16).setBackground('#e6f2ff'); // Actual Remaining (calculated with CF capacity)
    
    // Conditional formatting for ENTIRE PI Capacity Remaining (column 5)
    const remainingEntirePICell = sheet.getRange(currentRow, 5);
    if (remainingEntirePI < 0) {
      remainingEntirePICell.setBackground('#ffcccc').setFontColor('#cc0000').setFontWeight('bold');
    } else {
      remainingEntirePICell.setBackground('#d4edda').setFontColor('#155724').setFontWeight('bold');
    }
    
    // Conditional formatting for ENTIRE PI Planned Remaining (column 7)
    const anticipatedRemainingEntirePICell = sheet.getRange(currentRow, 7);
    if (anticipatedRemainingEntirePI < 0) {
      anticipatedRemainingEntirePICell.setBackground('#ffcccc').setFontColor('#cc0000').setFontWeight('bold');
    } else {
      anticipatedRemainingEntirePICell.setBackground('#d4edda').setFontColor('#155724').setFontWeight('bold');
    }
    
    // Conditional formatting for ENTIRE PI Actual Remaining (column 9)
    const actualRemainingEntirePICell = sheet.getRange(currentRow, 9);
    if (actualRemainingEntirePI < 0) {
      actualRemainingEntirePICell.setBackground('#ffcccc').setFontColor('#cc0000').setFontWeight('bold');
    } else {
      actualRemainingEntirePICell.setBackground('#d4edda').setFontColor('#155724').setFontWeight('bold');
    }
    
    // Conditional formatting for CODE FREEZE Capacity Remaining (column 12)
    // Uses blended colors to maintain blue tint while showing red/green status
    const remainingCodeFreezeCell = sheet.getRange(currentRow, 12);
    if (remainingCodeFreeze < 0) {
      remainingCodeFreezeCell.setBackground('#ffccdd').setFontColor('#cc0000').setFontWeight('bold'); // Pink-blue blend
    } else {
      remainingCodeFreezeCell.setBackground('#ccf2e6').setFontColor('#155724').setFontWeight('bold'); // Green-blue blend
    }
    
    // Conditional formatting for CODE FREEZE Planned Remaining (column 14)
    // Uses blended colors to maintain blue tint while showing red/green status
    const anticipatedRemainingCodeFreezeCell = sheet.getRange(currentRow, 14);
    if (anticipatedRemainingCodeFreeze < 0) {
      anticipatedRemainingCodeFreezeCell.setBackground('#ffccdd').setFontColor('#cc0000').setFontWeight('bold'); // Pink-blue blend
    } else {
      anticipatedRemainingCodeFreezeCell.setBackground('#ccf2e6').setFontColor('#155724').setFontWeight('bold'); // Green-blue blend
    }
    
    // Conditional formatting for CODE FREEZE Actual Remaining (column 16)
    // Uses blended colors to maintain blue tint while showing red/green status
    const actualRemainingCodeFreezeCell = sheet.getRange(currentRow, 16);
    if (actualRemainingCodeFreeze < 0) {
      actualRemainingCodeFreezeCell.setBackground('#ffccdd').setFontColor('#cc0000').setFontWeight('bold'); // Pink-blue blend
    } else {
      actualRemainingCodeFreezeCell.setBackground('#ccf2e6').setFontColor('#155724').setFontWeight('bold'); // Green-blue blend
    }
    
    currentRow++;
  }
  
  const dataEndRow = currentRow - 1;
  
  // Add TOTAL row (sum of SUBTOTAL + EYEFINITY only)
  sheet.getRange(currentRow, 1).setValue('TOTAL (Including EYEFINITY)');
  
  // If we have a subtotal row, sum SUBTOTAL + EYEFINITY rows only
  // Otherwise fall back to summing all data rows (shouldn't happen with EYEFINITY in config)
  if (subtotalRowNumber > 0 && eyefinityRowNumber > 0) {
    // Sum only the SUBTOTAL row and EYEFINITY row (rows 19 and 20 in the output)
    sheet.getRange(currentRow, 2).setFormula(`=B${subtotalRowNumber}+B${eyefinityRowNumber}`);
    sheet.getRange(currentRow, 3).setFormula(`=C${subtotalRowNumber}+C${eyefinityRowNumber}`);
    sheet.getRange(currentRow, 4).setFormula(`=D${subtotalRowNumber}+D${eyefinityRowNumber}`);
    sheet.getRange(currentRow, 5).setFormula(`=E${subtotalRowNumber}+E${eyefinityRowNumber}`);
    sheet.getRange(currentRow, 6).setFormula(`=F${subtotalRowNumber}+F${eyefinityRowNumber}`);
    sheet.getRange(currentRow, 7).setFormula(`=G${subtotalRowNumber}+G${eyefinityRowNumber}`);
    sheet.getRange(currentRow, 8).setFormula(`=H${subtotalRowNumber}+H${eyefinityRowNumber}`);
    sheet.getRange(currentRow, 9).setFormula(`=I${subtotalRowNumber}+I${eyefinityRowNumber}`);
    sheet.getRange(currentRow, 10).setFormula(`=J${subtotalRowNumber}+J${eyefinityRowNumber}`);
    sheet.getRange(currentRow, 11).setFormula(`=K${subtotalRowNumber}+K${eyefinityRowNumber}`);
    sheet.getRange(currentRow, 12).setFormula(`=L${subtotalRowNumber}+L${eyefinityRowNumber}`);
    sheet.getRange(currentRow, 13).setFormula(`=M${subtotalRowNumber}+M${eyefinityRowNumber}`);
    sheet.getRange(currentRow, 14).setFormula(`=N${subtotalRowNumber}+N${eyefinityRowNumber}`);
    sheet.getRange(currentRow, 15).setFormula(`=O${subtotalRowNumber}+O${eyefinityRowNumber}`);
    sheet.getRange(currentRow, 16).setFormula(`=P${subtotalRowNumber}+P${eyefinityRowNumber}`);
  } else {
    // Fallback: sum all data rows (old behavior)
    sheet.getRange(currentRow, 2).setFormula(`=SUM(B${dataStartRow}:B${dataEndRow})`);
    sheet.getRange(currentRow, 3).setFormula(`=SUM(C${dataStartRow}:C${dataEndRow})`);
    sheet.getRange(currentRow, 4).setFormula(`=SUM(D${dataStartRow}:D${dataEndRow})`);
    sheet.getRange(currentRow, 5).setFormula(`=SUM(E${dataStartRow}:E${dataEndRow})`);
    sheet.getRange(currentRow, 6).setFormula(`=SUM(F${dataStartRow}:F${dataEndRow})`);
    sheet.getRange(currentRow, 7).setFormula(`=SUM(G${dataStartRow}:G${dataEndRow})`);
    sheet.getRange(currentRow, 8).setFormula(`=SUM(H${dataStartRow}:H${dataEndRow})`);
    sheet.getRange(currentRow, 9).setFormula(`=SUM(I${dataStartRow}:I${dataEndRow})`);
    sheet.getRange(currentRow, 10).setFormula(`=SUM(J${dataStartRow}:J${dataEndRow})`);
    sheet.getRange(currentRow, 11).setFormula(`=SUM(K${dataStartRow}:K${dataEndRow})`);
    sheet.getRange(currentRow, 12).setFormula(`=SUM(L${dataStartRow}:L${dataEndRow})`);
    sheet.getRange(currentRow, 13).setFormula(`=SUM(M${dataStartRow}:M${dataEndRow})`);
    sheet.getRange(currentRow, 14).setFormula(`=SUM(N${dataStartRow}:N${dataEndRow})`);
    sheet.getRange(currentRow, 15).setFormula(`=SUM(O${dataStartRow}:O${dataEndRow})`);
    sheet.getRange(currentRow, 16).setFormula(`=SUM(P${dataStartRow}:P${dataEndRow})`);
  }
  
  // Format totals row
  sheet.getRange(currentRow, 1, 1, numColumns)
    .setFontWeight('bold')
    .setFontFamily('Comfortaa')
    .setBackground('#f0f0f0');
  
  sheet.getRange(currentRow, 2, 1, 15)
    .setHorizontalAlignment('center')
    .setNumberFormat('#,##0');
  
  // Keep Baseline Capacity light grey even in totals row
  sheet.getRange(currentRow, 2).setBackground('#d9d9d9'); // Medium grey for totals
  
  // Apply light blue tint to CODE FREEZE columns with different data
  sheet.getRange(currentRow, 10).setBackground('#c5dff0'); // Product Capacity - darker blue for totals
  sheet.getRange(currentRow, 12).setBackground('#c5dff0'); // Capacity Remaining (will be overridden by conditional)
  sheet.getRange(currentRow, 14).setBackground('#c5dff0'); // Planned Remaining (will be overridden by conditional)
  sheet.getRange(currentRow, 16).setBackground('#c5dff0'); // Actual Remaining (will be overridden by conditional)
  
  // Conditional formatting for ENTIRE PI Capacity Remaining (column 5)
  const totalRemainingEntirePICell = sheet.getRange(currentRow, 5);
  const totalRemainingEntirePI = totalRemainingEntirePICell.getValue();
  if (totalRemainingEntirePI < 0) {
    totalRemainingEntirePICell.setBackground('#ffcccc').setFontColor('#cc0000');
  } else {
    totalRemainingEntirePICell.setBackground('#d4edda').setFontColor('#155724');
  }
  
  // Conditional formatting for ENTIRE PI Planned Remaining (column 7)
  const totalAnticipatedRemainingEntirePICell = sheet.getRange(currentRow, 7);
  const totalAnticipatedRemainingEntirePI = totalAnticipatedRemainingEntirePICell.getValue();
  if (totalAnticipatedRemainingEntirePI < 0) {
    totalAnticipatedRemainingEntirePICell.setBackground('#ffcccc').setFontColor('#cc0000');
  } else {
    totalAnticipatedRemainingEntirePICell.setBackground('#d4edda').setFontColor('#155724');
  }
  
  // Conditional formatting for ENTIRE PI Actual Remaining (column 9)
  const totalActualRemainingEntirePICell = sheet.getRange(currentRow, 9);
  const totalActualRemainingEntirePI = totalActualRemainingEntirePICell.getValue();
  if (totalActualRemainingEntirePI < 0) {
    totalActualRemainingEntirePICell.setBackground('#ffcccc').setFontColor('#cc0000');
  } else {
    totalActualRemainingEntirePICell.setBackground('#d4edda').setFontColor('#155724');
  }
  
  // Conditional formatting for CODE FREEZE Capacity Remaining (column 12)
  // Uses blended colors to maintain blue tint
  const totalRemainingCodeFreezeCell = sheet.getRange(currentRow, 12);
  const totalRemainingCodeFreeze = totalRemainingCodeFreezeCell.getValue();
  if (totalRemainingCodeFreeze < 0) {
    totalRemainingCodeFreezeCell.setBackground('#ffccdd').setFontColor('#cc0000'); // Pink-blue blend
  } else {
    totalRemainingCodeFreezeCell.setBackground('#ccf2e6').setFontColor('#155724'); // Green-blue blend
  }
  
  // Conditional formatting for CODE FREEZE Planned Remaining (column 14)
  // Uses blended colors to maintain blue tint
  const totalAnticipatedRemainingCodeFreezeCell = sheet.getRange(currentRow, 14);
  const totalAnticipatedRemainingCodeFreeze = totalAnticipatedRemainingCodeFreezeCell.getValue();
  if (totalAnticipatedRemainingCodeFreeze < 0) {
    totalAnticipatedRemainingCodeFreezeCell.setBackground('#ffccdd').setFontColor('#cc0000'); // Pink-blue blend
  } else {
    totalAnticipatedRemainingCodeFreezeCell.setBackground('#ccf2e6').setFontColor('#155724'); // Green-blue blend
  }
  
  // Conditional formatting for CODE FREEZE Actual Remaining (column 16)
  // Uses blended colors to maintain blue tint
  const totalActualRemainingCodeFreezeCell = sheet.getRange(currentRow, 16);
  const totalActualRemainingCodeFreeze = totalActualRemainingCodeFreezeCell.getValue();
  if (totalActualRemainingCodeFreeze < 0) {
    totalActualRemainingCodeFreezeCell.setBackground('#ffccdd').setFontColor('#cc0000'); // Pink-blue blend
  } else {
    totalActualRemainingCodeFreezeCell.setBackground('#ccf2e6').setFontColor('#155724'); // Green-blue blend
  }
  
  currentRow++;
  
  // Add borders around the entire table (including category header row)
  const tableRange = sheet.getRange(headerRow - 1, 1, currentRow - (headerRow - 1), numColumns);
  tableRange.setBorder(
    true, true, true, true, true, true,
    'black', SpreadsheetApp.BorderStyle.SOLID
  );
  
  // Add thicker borders around paired variance columns to visually group them
  const tableHeight = currentRow - (headerRow - 1);
  
  // ENTIRE PI section pairs
  // Capacity Used (LOE) + Capacity Remaining (columns 4-5)
  sheet.getRange(headerRow - 1, 4, tableHeight, 2).setBorder(
    true, true, true, true, false, false,
    'black', SpreadsheetApp.BorderStyle.SOLID_MEDIUM
  );
  
  // Planned Load (FP) + Planned Remaining (columns 6-7)
  sheet.getRange(headerRow - 1, 6, tableHeight, 2).setBorder(
    true, true, true, true, false, false,
    'black', SpreadsheetApp.BorderStyle.SOLID_MEDIUM
  );
  
  // Actual Load (SPE) + Actual Remaining (columns 8-9)
  sheet.getRange(headerRow - 1, 8, tableHeight, 2).setBorder(
    true, true, true, true, false, false,
    'black', SpreadsheetApp.BorderStyle.SOLID_MEDIUM
  );
  
  // CODE FREEZE section pairs
  // Capacity Used (LOE) + Capacity Remaining (columns 11-12)
  sheet.getRange(headerRow - 1, 11, tableHeight, 2).setBorder(
    true, true, true, true, false, false,
    'black', SpreadsheetApp.BorderStyle.SOLID_MEDIUM
  );
  
  // Planned Load (FP) + Planned Remaining (columns 13-14)
  sheet.getRange(headerRow - 1, 13, tableHeight, 2).setBorder(
    true, true, true, true, false, false,
    'black', SpreadsheetApp.BorderStyle.SOLID_MEDIUM
  );
  
  // Actual Load (SPE) + Actual Remaining (columns 15-16)
  sheet.getRange(headerRow - 1, 15, tableHeight, 2).setBorder(
    true, true, true, true, false, false,
    'black', SpreadsheetApp.BorderStyle.SOLID_MEDIUM
  );
  
  return currentRow;
}

/**
 * Create table showing epics with blank Fix Version
 * @param {Sheet} sheet - The report sheet
 * @param {number} currentRow - Starting row for this table
 * @param {Sheet} piSheet - The PI data sheet to read from
 * @return {number} Next available row after the table
 */
function createBlankFixVersionTable(sheet, currentRow, piSheet) {
  if (!piSheet) {
    console.warn('PI sheet not provided - cannot create blank fix version table');
    return currentRow;
  }
  
  console.log('\n=== Creating Blank Fix Version Table ===');
  
  const dataRange = piSheet.getDataRange();
  const values = dataRange.getValues();
  
  if (values.length < 4) {
    console.warn('PI sheet has insufficient data');
    return currentRow;
  }
  
  // Headers are in row 4 (index 3)
  const headers = values[3];
  
  // Find column indices
  const keyCol = headers.indexOf('Key');
  const summaryCol = headers.indexOf('Summary');
  const scrumTeamCol = headers.indexOf('Scrum Team');
  const issueTypeCol = headers.indexOf('Issue Type');
  const fixVersionCol = headers.indexOf('Fix Version');
  const featurePointsCol = headers.indexOf('Feature Points');
  const storyPointEstimateCol = headers.indexOf('Story Point Estimate');
  
  if (keyCol === -1 || scrumTeamCol === -1 || issueTypeCol === -1 || fixVersionCol === -1) {
    console.error('Required columns not found in PI sheet');
    return currentRow;
  }
  
  // Collect epics with blank fix version for clinical teams
  const blankFixVersionEpics = [];
  
  for (let i = 4; i < values.length; i++) {
    const row = values[i];
    const key = row[keyCol];
    const summary = summaryCol !== -1 ? row[summaryCol] : '';
    const scrumTeam = row[scrumTeamCol];
    const issueType = row[issueTypeCol];
    const fixVersion = row[fixVersionCol];
    const featurePoints = featurePointsCol !== -1 ? (parseFloat(row[featurePointsCol]) || 0) : 0;
    const storyPointEstimate = storyPointEstimateCol !== -1 ? (parseFloat(row[storyPointEstimateCol]) || 0) : 0;
    
    // Must be an Epic
    if (issueType !== 'Epic') {
      continue;
    }
    
    // Must have blank/empty Fix Version
    if (fixVersion && fixVersion.toString().trim() !== '') {
      continue;
    }
    
    // Must be a clinical team
    const cleanTeamName = scrumTeam ? scrumTeam.toString().trim().toUpperCase() : '';
    if (!DANS_REPORT_CONFIG.clinicalTeams.includes(cleanTeamName)) {
      continue;
    }
    
    blankFixVersionEpics.push({
      key: key,
      summary: summary,
      scrumTeam: scrumTeam,
      featurePoints: Math.ceil(featurePoints), // Round up to whole number
      storyPointEstimate: Math.ceil(storyPointEstimate) // Round up to whole number
    });
  }
  
  console.log(`Found ${blankFixVersionEpics.length} epics with blank fix version`);
  
  // Sort epics by scrum team, with EYEFINITY at the bottom
  blankFixVersionEpics.sort((a, b) => {
    const teamA = a.scrumTeam.toString().trim().toUpperCase();
    const teamB = b.scrumTeam.toString().trim().toUpperCase();
    
    // EYEFINITY always goes to the bottom
    if (teamA === 'EYEFINITY' && teamB !== 'EYEFINITY') return 1;
    if (teamA !== 'EYEFINITY' && teamB === 'EYEFINITY') return -1;
    
    // Otherwise sort alphabetically
    return teamA.localeCompare(teamB);
  });
  
  console.log('Epics sorted by scrum team (EYEFINITY at bottom)');
  
  if (blankFixVersionEpics.length === 0) {
    // No epics found - don't create the table
    console.log('No epics with blank fix version found - skipping table');
    return currentRow;
  }
  
  // Add section title (now spans across 7 columns to match new table width)
  sheet.getRange(currentRow, 1).setValue('EPICS WITH BLANK FIX VERSION');
  sheet.getRange(currentRow, 1)
    .setFontSize(14)
    .setFontWeight('bold')
    .setFontFamily('Comfortaa')
    .setBackground('#ffcccc')
    .setFontColor('#cc0000');
  sheet.getRange(currentRow, 1, 1, 7).merge(); // Merge across 7 columns now
  currentRow++;
  
  // NEW TABLE STRUCTURE:
  // Column 1: Key
  // Columns 2-4: Summary (merged)
  // Column 5: Scrum Team
  // Column 6: Feature Point
  // Column 7: Story Point Estimate
  
  const headerRow = currentRow;
  
  // Create headers with merging
  sheet.getRange(currentRow, 1).setValue('Key');
  sheet.getRange(currentRow, 2, 1, 3).merge(); // Merge columns 2, 3, 4 for Summary
  sheet.getRange(currentRow, 2).setValue('Summary');
  sheet.getRange(currentRow, 5).setValue('Scrum Team');
  sheet.getRange(currentRow, 6).setValue('Feature Point');
  sheet.getRange(currentRow, 7).setValue('Story Point Estimate');
  
  // Format all headers
  sheet.getRange(currentRow, 1, 1, 7)
    .setFontWeight('bold')
    .setBackground('#9b7bb8')
    .setFontColor('white')
    .setFontSize(10)
    .setFontFamily('Comfortaa')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true);
  
  // Add notes to column headers
  sheet.getRange(currentRow, 5).setNote('Sorted alphabetically\nEYEFINITY appears at the bottom');
  sheet.getRange(currentRow, 6).setNote('Feature Points at Epic level\nUsed to calculate Planned Load (FP Ã— 10)');
  sheet.getRange(currentRow, 7).setNote('Story Point Estimate at Epic level\nUsed to calculate Actual Load');
  
  currentRow++;
  
  const dataStartRow = currentRow;
  
  // Add data rows
  for (const epic of blankFixVersionEpics) {
    // Column 1: Key (with hyperlink)
    const epicUrl = `${DANS_REPORT_CONFIG.jiraBaseUrl}/browse/${epic.key}`;
    sheet.getRange(currentRow, 1).setFormula(`=HYPERLINK("${epicUrl}", "${epic.key}")`);
    
    // Columns 2-4: Summary (merged)
    sheet.getRange(currentRow, 2, 1, 3).merge();
    sheet.getRange(currentRow, 2).setValue(epic.summary);
    sheet.getRange(currentRow, 2).setWrap(true); // Allow text to wrap
    
    // Column 5: Scrum Team
    sheet.getRange(currentRow, 5).setValue(epic.scrumTeam);
    
    // Column 6: Feature Point
    sheet.getRange(currentRow, 6).setValue(epic.featurePoints);
    
    // Column 7: Story Point Estimate
    sheet.getRange(currentRow, 7).setValue(epic.storyPointEstimate);
    
    // Format this row
    sheet.getRange(currentRow, 1, 1, 7)
      .setFontSize(9)
      .setFontFamily('Comfortaa')
      .setVerticalAlignment('middle');
    
    // CENTER JUSTIFY number columns (6 and 7)
    sheet.getRange(currentRow, 6, 1, 2)
      .setHorizontalAlignment('center')
      .setNumberFormat('#,##0');
    
    currentRow++;
  }
  
  const dataEndRow = currentRow - 1;
  
  // Add totals row
  sheet.getRange(currentRow, 1).setValue('TOTAL');
  sheet.getRange(currentRow, 2, 1, 3).merge(); // Merge summary columns for totals row
  sheet.getRange(currentRow, 2).setValue(''); // Empty merged cells
  sheet.getRange(currentRow, 5).setValue(''); // Empty scrum team for totals
  sheet.getRange(currentRow, 6).setFormula(`=SUM(F${dataStartRow}:F${dataEndRow})`); // Feature Point total
  sheet.getRange(currentRow, 7).setFormula(`=SUM(G${dataStartRow}:G${dataEndRow})`); // Story Point Estimate total
  
  // Format totals row
  sheet.getRange(currentRow, 1, 1, 7)
    .setFontWeight('bold')
    .setFontFamily('Comfortaa')
    .setBackground('#f0f0f0');
  
  // CENTER JUSTIFY number columns in totals (6 and 7)
  sheet.getRange(currentRow, 6, 1, 2)
    .setHorizontalAlignment('center')
    .setNumberFormat('#,##0');
  
  currentRow++;
  
  // Set column widths
  sheet.setColumnWidth(1, 100);  // Key
  sheet.setColumnWidth(2, 150);  // Summary part 1
  sheet.setColumnWidth(3, 150);  // Summary part 2
  sheet.setColumnWidth(4, 150);  // Summary part 3
  sheet.setColumnWidth(5, 120);  // Scrum Team
  sheet.setColumnWidth(6, 100);  // Feature Point
  sheet.setColumnWidth(7, 120);  // Story Point Estimate
  
  // Add borders around the entire table
  const tableRange = sheet.getRange(headerRow, 1, currentRow - headerRow, 7);
  tableRange.setBorder(
    true, true, true, true, true, true,
    'black', SpreadsheetApp.BorderStyle.SOLID
  );
  
  console.log(`Blank Fix Version table created with ${blankFixVersionEpics.length} epics`);
  
  return currentRow;
}

// ===== ROLE-BASED CAPACITY FUNCTIONS =====

/**
 * Detect role from a ticket's labels or title prefix
 * Uses label matching first, then falls back to title prefix patterns
 * 
 * @param {Array} labels - Array of label strings from the ticket
 * @param {string} summary - The ticket summary/title
 * @return {string|null} Normalized role key or null if not detected
 */
function detectRoleFromTicketDans(labels, summary) {
  // Priority 1: Check labels for role matches
  if (labels && labels.length > 0) {
    for (const label of labels) {
      const cleanLabel = label.toString().trim().toUpperCase();
      const normalized = ROLE_NORMALIZATION_DANS[cleanLabel];
      if (normalized) {
        return normalized;
      }
    }
  }
  
  // Priority 2: Check title prefix patterns like [BE], FE:, (QA), BE -, etc.
  if (summary) {
    const titleStr = summary.toString().trim();
    // Match patterns: [ROLE], (ROLE), ROLE:, ROLE -, ROLE_
    const prefixMatch = titleStr.match(/^[\[\(]?\s*(QA|AQA|W-DEV|WDEV|M-DEV|MDEV|BE|FE|MOBILE|DEVOPS|UX|M-ANDROID|M-IOS)\s*[\]\):_\-]/i);
    if (prefixMatch) {
      const roleKey = prefixMatch[1].toUpperCase();
      return ROLE_NORMALIZATION_DANS[roleKey] || roleKey;
    }
  }
  
  return null;
}

/**
 * Get role-level capacity data from the PI Capacity sheet for EMA Clinical teams
 * 
 * Reads the "PI{N} - Capacity" sheet and finds the EMA Clinical section.
 * For each clinical team, extracts Before FF and After FF role capacities.
 * 
 * Returns aggregated role capacity across all clinical teams:
 * - entirePI: Before FF + After FF totals per role
 * - codeFreeze: Before FF totals only per role (iterations 1-5)
 * 
 * @param {Spreadsheet} spreadsheet - The spreadsheet object
 * @param {string} piNumber - The PI number (e.g., "14")
 * @return {Object} { entirePI: { ROLE: capacity, ... }, codeFreeze: { ROLE: capacity, ... }, teamRoles: { TEAM: { ROLE: { beforeFF, afterFF, entirePI }, ... }, ... } }
 */
function getRoleCapacityData(spreadsheet, piNumber) {
  const capacitySheetName = `PI${piNumber} - Capacity`;
  const capacitySheet = spreadsheet.getSheetByName(capacitySheetName);
  
  if (!capacitySheet) {
    console.warn(`Role capacity sheet "${capacitySheetName}" not found`);
    return { entirePI: {}, codeFreeze: {}, teamRoles: {} };
  }
  
  console.log(`\n=== Reading Role Capacity Data from "${capacitySheetName}" ===`);
  
  const dataRange = capacitySheet.getDataRange();
  const values = dataRange.getValues();
  const maxRows = values.length;
  const maxCols = values[0] ? values[0].length : 0;
  
  // Find the EMA Clinical column by scanning row 1 for "EMA Clinical"
  let clinicalCol = -1;
  for (let col = 0; col < maxCols; col++) {
    const headerVal = values[0][col];
    if (headerVal && headerVal.toString().trim().toLowerCase().includes('ema clinical')) {
      clinicalCol = col;
      break;
    }
  }
  
  if (clinicalCol === -1) {
    console.warn('EMA Clinical section not found in capacity sheet row 1');
    return { entirePI: {}, codeFreeze: {}, teamRoles: {} };
  }
  
  console.log(`Found EMA Clinical at column ${clinicalCol + 1} (${String.fromCharCode(65 + (clinicalCol % 26))})`);
  
  // The Total column is at offset +9 from the team column
  const totalColOffset = 9;
  
  // Aggregated role capacity across all teams
  const entirePIRoles = {};
  const codeFreezeRoles = {};
  const teamRoles = {};
  
  // Scan for team blocks - team name is in the clinical column, followed by structured data
  let row = 2; // Start after header row
  while (row < maxRows) {
    const cellValue = values[row][clinicalCol];
    
    // Skip empty rows and known non-team markers
    if (!cellValue || cellValue.toString().trim() === '' || cellValue.toString().trim() === '-') {
      row++;
      continue;
    }
    
    const cellStr = cellValue.toString().trim();
    
    // Skip known structural markers
    if (cellStr.toLowerCase().includes('allocation type') ||
        cellStr.toLowerCase().includes('base capacity') ||
        cellStr.toLowerCase().includes('klo') ||
        cellStr.toLowerCase().includes('quality') ||
        cellStr.toLowerCase().includes('tech') ||
        cellStr.toLowerCase().includes('product') ||
        cellStr.toLowerCase().includes('unplanned')) {
      row++;
      continue;
    }
    
    // Check if next row is "Allocation Type" - confirms this is a team name
    if (row + 1 < maxRows) {
      const nextRowVal = values[row + 1][clinicalCol];
      if (nextRowVal && nextRowVal.toString().trim() === 'Allocation Type') {
        const teamName = cellStr.toUpperCase();
        
        // Check if this is a clinical team we care about
        const isClinicalTeam = DANS_REPORT_CONFIG.clinicalTeams.some(
          t => t === teamName || t === cellStr.toUpperCase().replace(/\s+/g, ' ')
        );
        
        if (!isClinicalTeam) {
          // Not a clinical team, skip this block
          row += 25; // Jump past this team block
          continue;
        }
        
        console.log(`Processing clinical team: ${teamName} at row ${row + 1}`);
        teamRoles[teamName] = {};
        
        // Find "Base Capacity before FF" within the next 12 rows
        let beforeFFRow = -1;
        for (let offset = 8; offset <= 12; offset++) {
          if (row + offset < maxRows) {
            const checkVal = values[row + offset][clinicalCol];
            if (checkVal && checkVal.toString().toLowerCase().includes('before ff')) {
              beforeFFRow = row + offset;
              break;
            }
          }
        }
        
        // Find "Base Capacity after FF" within the next 20 rows
        let afterFFRow = -1;
        for (let offset = 16; offset <= 20; offset++) {
          if (row + offset < maxRows) {
            const checkVal = values[row + offset][clinicalCol];
            if (checkVal && checkVal.toString().toLowerCase().includes('after ff')) {
              afterFFRow = row + offset;
              break;
            }
          }
        }
        
        // Read Before FF roles (rows after beforeFFRow header)
        if (beforeFFRow > 0) {
          for (let rOffset = 1; rOffset <= 6; rOffset++) {
            const roleRow = beforeFFRow + rOffset;
            if (roleRow >= maxRows) break;
            
            const roleName = values[roleRow][clinicalCol];
            if (!roleName || roleName.toString().trim() === '' || roleName.toString().trim() === '-') continue;
            
            // Stop if we hit a subtotal row (no role name, just numbers)
            const roleStr = roleName.toString().trim();
            if (roleStr.toLowerCase().includes('base capacity')) break;
            
            const roleKey = roleStr.toUpperCase();
            const normalizedRole = ROLE_NORMALIZATION_DANS[roleKey] || roleKey;
            
            // Get the total value (at totalColOffset from clinicalCol)
            const totalVal = values[roleRow][clinicalCol + totalColOffset];
            let beforeFFTotal = 0;
            if (totalVal && totalVal !== '-' && totalVal !== '') {
              beforeFFTotal = parseFloat(totalVal) || 0;
            }
            
            if (!teamRoles[teamName][normalizedRole]) {
              teamRoles[teamName][normalizedRole] = { beforeFF: 0, afterFF: 0, entirePI: 0 };
            }
            teamRoles[teamName][normalizedRole].beforeFF = Math.ceil(beforeFFTotal);
          }
        }
        
        // Read After FF roles (rows after afterFFRow header)
        if (afterFFRow > 0) {
          for (let rOffset = 1; rOffset <= 6; rOffset++) {
            const roleRow = afterFFRow + rOffset;
            if (roleRow >= maxRows) break;
            
            const roleName = values[roleRow][clinicalCol];
            if (!roleName || roleName.toString().trim() === '' || roleName.toString().trim() === '-') continue;
            
            const roleStr = roleName.toString().trim();
            if (roleStr.toLowerCase().includes('base capacity')) break;
            
            const roleKey = roleStr.toUpperCase();
            const normalizedRole = ROLE_NORMALIZATION_DANS[roleKey] || roleKey;
            
            const totalVal = values[roleRow][clinicalCol + totalColOffset];
            let afterFFTotal = 0;
            if (totalVal && totalVal !== '-' && totalVal !== '') {
              afterFFTotal = parseFloat(totalVal) || 0;
            }
            
            if (!teamRoles[teamName][normalizedRole]) {
              teamRoles[teamName][normalizedRole] = { beforeFF: 0, afterFF: 0, entirePI: 0 };
            }
            teamRoles[teamName][normalizedRole].afterFF = Math.ceil(afterFFTotal);
          }
        }
        
        // Calculate Entire PI totals and aggregate across teams
        Object.keys(teamRoles[teamName]).forEach(role => {
          const rd = teamRoles[teamName][role];
          rd.entirePI = rd.beforeFF + rd.afterFF;
          
          if (!entirePIRoles[role]) entirePIRoles[role] = 0;
          if (!codeFreezeRoles[role]) codeFreezeRoles[role] = 0;
          
          entirePIRoles[role] += rd.entirePI;
          codeFreezeRoles[role] += rd.beforeFF; // Code Freeze = Before FF only
        });
        
        console.log(`  Roles found: ${Object.keys(teamRoles[teamName]).filter(r => teamRoles[teamName][r].entirePI > 0).join(', ')}`);
        
        // Move past this team block
        row += 25;
        continue;
      }
    }
    
    row++;
  }
  
  // Log summary
  console.log('\n=== Aggregated Role Capacity Summary ===');
  Object.keys(entirePIRoles).sort().forEach(role => {
    if (entirePIRoles[role] > 0 || codeFreezeRoles[role] > 0) {
      console.log(`  ${role}: Entire PI = ${entirePIRoles[role]}, Code Freeze = ${codeFreezeRoles[role]}`);
    }
  });
  console.log('=== End Role Capacity ===\n');
  
  return { entirePI: entirePIRoles, codeFreeze: codeFreezeRoles, teamRoles: teamRoles };
}

/**
 * Get capacity used (story points) by ROLE from the PI data sheet
 * Uses role detection (labels + title prefix) to assign story points to roles
 * 
 * @param {Sheet} piSheet - The PI data sheet
 * @return {Object} { ROLE: totalStoryPoints, ... }
 */
function getCapacityUsedByRole(piSheet) {
  if (!piSheet) {
    console.warn('PI sheet not provided - role capacity used will be 0');
    return {};
  }
  
  console.log(`\n=== Calculating Capacity Used by Role ===`);
  
  const dataRange = piSheet.getDataRange();
  const values = dataRange.getValues();
  
  if (values.length < 4) return {};
  
  const headers = values[3];
  const scrumTeamCol = headers.indexOf('Scrum Team');
  const storyPointsCol = headers.indexOf('Story Points');
  const issueTypeCol = headers.indexOf('Issue Type');
  const allocationCol = headers.indexOf('Allocation');
  const labelsCol = headers.indexOf('Labels');
  const summaryCol = headers.indexOf('Summary');
  
  if (scrumTeamCol === -1 || storyPointsCol === -1) {
    console.error('Required columns not found for role capacity used');
    return {};
  }
  
  const roleUsed = {};
  let matched = 0;
  let unmatched = 0;
  
  for (let i = 4; i < values.length; i++) {
    const row = values[i];
    const scrumTeam = row[scrumTeamCol];
    const storyPoints = parseFloat(row[storyPointsCol]) || 0;
    const issueType = issueTypeCol !== -1 ? row[issueTypeCol] : '';
    const allocation = allocationCol !== -1 ? row[allocationCol] : '';
    
    if (!scrumTeam || scrumTeam.toString().trim() === '') continue;
    if (issueType !== 'Story' && issueType !== 'Bug') continue;
    if (allocation !== 'Product - Feature' && allocation !== 'Product - Compliance') continue;
    if (storyPoints === 0) continue;
    
    // Only count clinical teams
    const cleanTeamName = scrumTeam.toString().trim().toUpperCase();
    if (!DANS_REPORT_CONFIG.clinicalTeams.includes(cleanTeamName)) continue;
    
    // Detect role
    const labels = labelsCol !== -1 ? (row[labelsCol] || '').toString().split(',').map(l => l.trim()) : [];
    const summary = summaryCol !== -1 ? row[summaryCol] : '';
    const detectedRole = detectRoleFromTicketDans(labels, summary) || 'Unassigned';
    
    if (!roleUsed[detectedRole]) roleUsed[detectedRole] = 0;
    roleUsed[detectedRole] += Math.ceil(storyPoints);
    
    if (detectedRole === 'Unassigned') unmatched++;
    else matched++;
  }
  
  console.log(`Role detection: ${matched} matched, ${unmatched} unassigned`);
  console.log('Role Used:', JSON.stringify(roleUsed));
  console.log('=== End Capacity Used by Role ===\n');
  
  return roleUsed;
}

/**
 * Get load data (Planned Load FP and Actual Load SPE) by ROLE from Epic-level data
 * 
 * For Epics, role is determined by looking at the Epic's own labels/title prefix.
 * If an Epic has no detectable role, its load is assigned to 'Unassigned'.
 * 
 * @param {Sheet} piSheet - The PI data sheet
 * @return {Object} { anticipatedLoad: { ROLE: value }, actualLoad: { ROLE: value } }
 */
function getLoadDataByRole(piSheet) {
  if (!piSheet) {
    return { anticipatedLoad: {}, actualLoad: {} };
  }
  
  console.log(`\n=== Calculating Load Data by Role ===`);
  
  const dataRange = piSheet.getDataRange();
  const values = dataRange.getValues();
  
  if (values.length < 4) return { anticipatedLoad: {}, actualLoad: {} };
  
  const headers = values[3];
  const scrumTeamCol = headers.indexOf('Scrum Team');
  const issueTypeCol = headers.indexOf('Issue Type');
  const allocationCol = headers.indexOf('Allocation');
  const featurePointsCol = headers.indexOf('Feature Points');
  const storyPointEstimateCol = headers.indexOf('Story Point Estimate');
  const labelsCol = headers.indexOf('Labels');
  const summaryCol = headers.indexOf('Summary');
  
  if (scrumTeamCol === -1 || issueTypeCol === -1 || allocationCol === -1) {
    return { anticipatedLoad: {}, actualLoad: {} };
  }
  
  const anticipatedLoad = {};
  const actualLoad = {};
  
  for (let i = 4; i < values.length; i++) {
    const row = values[i];
    const scrumTeam = row[scrumTeamCol];
    const issueType = row[issueTypeCol];
    const allocation = row[allocationCol];
    
    if (!scrumTeam || scrumTeam.toString().trim() === '') continue;
    if (issueType !== 'Epic') continue;
    if (allocation !== 'Product - Feature' && allocation !== 'Product - Compliance') continue;
    
    const cleanTeamName = scrumTeam.toString().trim().toUpperCase();
    if (!DANS_REPORT_CONFIG.clinicalTeams.includes(cleanTeamName)) continue;
    
    // Detect role from Epic labels/title
    const labels = labelsCol !== -1 ? (row[labelsCol] || '').toString().split(',').map(l => l.trim()) : [];
    const summary = summaryCol !== -1 ? row[summaryCol] : '';
    const detectedRole = detectRoleFromTicketDans(labels, summary) || 'Unassigned';
    
    const featurePoints = featurePointsCol !== -1 ? parseFloat(row[featurePointsCol]) || 0 : 0;
    const storyPointEstimate = storyPointEstimateCol !== -1 ? parseFloat(row[storyPointEstimateCol]) || 0 : 0;
    
    if (!anticipatedLoad[detectedRole]) anticipatedLoad[detectedRole] = 0;
    if (!actualLoad[detectedRole]) actualLoad[detectedRole] = 0;
    
    if (featurePoints > 0) {
      anticipatedLoad[detectedRole] += Math.ceil(featurePoints) * 10;
    }
    if (storyPointEstimate > 0) {
      actualLoad[detectedRole] += Math.ceil(storyPointEstimate);
    }
  }
  
  console.log('Planned Load by Role (FP×10):', JSON.stringify(anticipatedLoad));
  console.log('Actual Load by Role (SPE):', JSON.stringify(actualLoad));
  console.log('=== End Load Data by Role ===\n');
  
  return { anticipatedLoad, actualLoad };
}

/**
 * Create the role-based capacity utilization table
 * Same 16-column structure as the team-based table but with roles as rows
 * 
 * @param {Sheet} sheet - The report sheet
 * @param {number} currentRow - Starting row for this table
 * @param {Object} roleCapacity - From getRoleCapacityData()
 * @param {Object} roleUsed - From getCapacityUsedByRole()
 * @param {Object} roleLoadData - From getLoadDataByRole()
 * @return {number} Next available row after the table
 */
function createRoleCapacityTable(sheet, currentRow, roleCapacity, roleUsed, roleLoadData) {
  const numColumns = 16;
  
  // Section title
  sheet.getRange(currentRow, 1).setValue('CLINICAL CAPACITY UTILIZATION BY ROLE - ENTIRE PI vs CODE FREEZE');
  sheet.getRange(currentRow, 1)
    .setFontSize(14)
    .setFontWeight('bold')
    .setFontFamily('Comfortaa')
    .setBackground('#E1D5E7')
    .setFontColor('black');
  sheet.getRange(currentRow, 1, 1, numColumns).merge();
  currentRow++;
  
  // Category headers
  sheet.getRange(currentRow, 1).setValue('');
  sheet.getRange(currentRow, 2).setValue('');
  sheet.getRange(currentRow, 3, 1, 7).merge();
  sheet.getRange(currentRow, 3).setValue('ENTIRE PI');
  sheet.getRange(currentRow, 10, 1, 7).merge();
  sheet.getRange(currentRow, 10).setValue('CODE FREEZE');
  
  sheet.getRange(currentRow, 1, 1, numColumns)
    .setFontWeight('bold')
    .setBackground('#D5A6E0')
    .setFontColor('black')
    .setFontSize(10)
    .setFontFamily('Comfortaa')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  currentRow++;
  
  // Column headers (same as team table but "Role" instead of "Scrum Team")
  const headers = [
    'Role',
    'Sprint Capacity',
    'Baseline Capacity',
    'Capacity Used (LOE)',
    'Capacity Remaining',
    'Planned Load (FP)',
    'Planned Remaining',
    'Actual Load (SPE)',
    'Actual Remaining',
    'Baseline Capacity',
    'Capacity Used (LOE)',
    'Capacity Remaining',
    'Planned Load (FP)',
    'Planned Remaining',
    'Actual Load (SPE)',
    'Actual Remaining'
  ];
  
  sheet.getRange(currentRow, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(currentRow, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#9b7bb8')
    .setFontColor('white')
    .setFontSize(9)
    .setFontFamily('Comfortaa')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true);
  
  // Add notes to role table headers
  sheet.getRange(currentRow, 2).setNote('Sum of Before FF + After FF capacity for this role across all clinical teams');
  sheet.getRange(currentRow, 3).setNote('ENTIRE PI: Before FF + After FF capacity for this role\nAggregated across all clinical teams');
  sheet.getRange(currentRow, 4).setNote('Story points from Stories/Bugs assigned to this role\n(detected from labels or title prefix)');
  sheet.getRange(currentRow, 10).setNote('CODE FREEZE: Before FF capacity only (iterations 1-5)\nLIGHT BLUE = Different data source');
  
  const headerRow = currentRow;
  currentRow++;
  
  // Determine which roles to display
  // Collect all roles from capacity and usage, sorted with Unassigned last
  const allRoles = new Set();
  Object.keys(roleCapacity.entirePI).forEach(r => { if (roleCapacity.entirePI[r] > 0) allRoles.add(r); });
  Object.keys(roleCapacity.codeFreeze).forEach(r => { if (roleCapacity.codeFreeze[r] > 0) allRoles.add(r); });
  Object.keys(roleUsed).forEach(r => allRoles.add(r));
  Object.keys(roleLoadData.anticipatedLoad).forEach(r => allRoles.add(r));
  Object.keys(roleLoadData.actualLoad).forEach(r => allRoles.add(r));
  
  // Remove Unassigned from main list (will be added at end)
  allRoles.delete('Unassigned');
  
  // Sort roles alphabetically
  const sortedRoles = Array.from(allRoles).sort();
  
  // Add Unassigned at the end if it has any data
  const hasUnassigned = (roleUsed['Unassigned'] || 0) > 0 || 
                        (roleLoadData.anticipatedLoad['Unassigned'] || 0) > 0 ||
                        (roleLoadData.actualLoad['Unassigned'] || 0) > 0;
  if (hasUnassigned) {
    sortedRoles.push('Unassigned');
  }
  
  const dataStartRow = currentRow;
  
  // Write data rows for each role
  for (const role of sortedRoles) {
    const isUnassigned = role === 'Unassigned';
    
    // Capacity values
    const entirePICap = roleCapacity.entirePI[role] || 0;
    const codeFreezeCap = roleCapacity.codeFreeze[role] || 0;
    const capacityUsed = roleUsed[role] || 0;
    
    // Sprint Capacity = Entire PI capacity (same concept as team table)
    const sprintCapacity = entirePICap;
    
    // Remaining
    const remainingEntirePI = entirePICap - capacityUsed;
    const remainingCodeFreeze = codeFreezeCap - capacityUsed;
    
    // Load data
    const anticipatedLoad = roleLoadData.anticipatedLoad[role] || 0;
    const actualLoad = roleLoadData.actualLoad[role] || 0;
    
    // Planned/Actual Remaining
    const anticipatedRemainingEntirePI = entirePICap - anticipatedLoad;
    const actualRemainingEntirePI = entirePICap - actualLoad;
    const anticipatedRemainingCodeFreeze = codeFreezeCap - anticipatedLoad;
    const actualRemainingCodeFreeze = codeFreezeCap - actualLoad;
    
    // Write row
    sheet.getRange(currentRow, 1).setValue(role);
    sheet.getRange(currentRow, 2).setValue(sprintCapacity);
    
    // ENTIRE PI section
    sheet.getRange(currentRow, 3).setValue(entirePICap);
    sheet.getRange(currentRow, 4).setValue(capacityUsed);
    sheet.getRange(currentRow, 5).setValue(remainingEntirePI);
    sheet.getRange(currentRow, 6).setValue(anticipatedLoad);
    sheet.getRange(currentRow, 7).setValue(anticipatedRemainingEntirePI);
    sheet.getRange(currentRow, 8).setValue(actualLoad);
    sheet.getRange(currentRow, 9).setValue(actualRemainingEntirePI);
    
    // CODE FREEZE section
    sheet.getRange(currentRow, 10).setValue(codeFreezeCap);
    sheet.getRange(currentRow, 11).setValue(capacityUsed);
    sheet.getRange(currentRow, 12).setValue(remainingCodeFreeze);
    sheet.getRange(currentRow, 13).setValue(anticipatedLoad);
    sheet.getRange(currentRow, 14).setValue(anticipatedRemainingCodeFreeze);
    sheet.getRange(currentRow, 15).setValue(actualLoad);
    sheet.getRange(currentRow, 16).setValue(actualRemainingCodeFreeze);
    
    // Format row
    sheet.getRange(currentRow, 1, 1, numColumns)
      .setFontSize(9)
      .setFontFamily('Comfortaa')
      .setVerticalAlignment('middle');
    
    sheet.getRange(currentRow, 2, 1, 15)
      .setHorizontalAlignment('center')
      .setNumberFormat('#,##0');
    
    // Light grey for Sprint Capacity
    sheet.getRange(currentRow, 2).setBackground('#f5f5f5');
    
    // Light blue for CODE FREEZE columns with different data source
    sheet.getRange(currentRow, 10).setBackground('#e6f2ff');
    sheet.getRange(currentRow, 12).setBackground('#e6f2ff');
    sheet.getRange(currentRow, 14).setBackground('#e6f2ff');
    sheet.getRange(currentRow, 16).setBackground('#e6f2ff');
    
    // Yellow highlight for Unassigned
    if (isUnassigned) {
      sheet.getRange(currentRow, 1).setBackground('#fff3cd');
      if (capacityUsed > 0) {
        sheet.getRange(currentRow, 4).setBackground('#fff3cd');
        sheet.getRange(currentRow, 11).setBackground('#fff3cd');
      }
    }
    
    // Conditional formatting for remaining columns - ENTIRE PI
    const remainEntirePICell = sheet.getRange(currentRow, 5);
    if (remainingEntirePI < 0) {
      remainEntirePICell.setBackground('#ffcccc').setFontColor('#cc0000').setFontWeight('bold');
    } else {
      remainEntirePICell.setBackground('#d4edda').setFontColor('#155724').setFontWeight('bold');
    }
    
    const anticipatedRemEntirePICell = sheet.getRange(currentRow, 7);
    if (anticipatedRemainingEntirePI < 0) {
      anticipatedRemEntirePICell.setBackground('#ffcccc').setFontColor('#cc0000').setFontWeight('bold');
    } else {
      anticipatedRemEntirePICell.setBackground('#d4edda').setFontColor('#155724').setFontWeight('bold');
    }
    
    const actualRemEntirePICell = sheet.getRange(currentRow, 9);
    if (actualRemainingEntirePI < 0) {
      actualRemEntirePICell.setBackground('#ffcccc').setFontColor('#cc0000').setFontWeight('bold');
    } else {
      actualRemEntirePICell.setBackground('#d4edda').setFontColor('#155724').setFontWeight('bold');
    }
    
    // Conditional formatting for remaining columns - CODE FREEZE (blended colors)
    const remainCFCell = sheet.getRange(currentRow, 12);
    if (remainingCodeFreeze < 0) {
      remainCFCell.setBackground('#ffccdd').setFontColor('#cc0000').setFontWeight('bold');
    } else {
      remainCFCell.setBackground('#ccf2e6').setFontColor('#155724').setFontWeight('bold');
    }
    
    const anticipatedRemCFCell = sheet.getRange(currentRow, 14);
    if (anticipatedRemainingCodeFreeze < 0) {
      anticipatedRemCFCell.setBackground('#ffccdd').setFontColor('#cc0000').setFontWeight('bold');
    } else {
      anticipatedRemCFCell.setBackground('#ccf2e6').setFontColor('#155724').setFontWeight('bold');
    }
    
    const actualRemCFCell = sheet.getRange(currentRow, 16);
    if (actualRemainingCodeFreeze < 0) {
      actualRemCFCell.setBackground('#ffccdd').setFontColor('#cc0000').setFontWeight('bold');
    } else {
      actualRemCFCell.setBackground('#ccf2e6').setFontColor('#155724').setFontWeight('bold');
    }
    
    currentRow++;
  }
  
  const dataEndRow = currentRow - 1;
  
  // Add TOTAL row
  sheet.getRange(currentRow, 1).setValue('TOTAL');
  
  for (let col = 2; col <= 16; col++) {
    const colLetter = String.fromCharCode(64 + col); // B=66, C=67, etc.
    sheet.getRange(currentRow, col).setFormula(`=SUM(${colLetter}${dataStartRow}:${colLetter}${dataEndRow})`);
  }
  
  // Format totals row
  sheet.getRange(currentRow, 1, 1, numColumns)
    .setFontWeight('bold')
    .setFontFamily('Comfortaa')
    .setBackground('#f0f0f0');
  
  sheet.getRange(currentRow, 2, 1, 15)
    .setHorizontalAlignment('center')
    .setNumberFormat('#,##0');
  
  sheet.getRange(currentRow, 2).setBackground('#d9d9d9');
  sheet.getRange(currentRow, 10).setBackground('#c5dff0');
  sheet.getRange(currentRow, 12).setBackground('#c5dff0');
  sheet.getRange(currentRow, 14).setBackground('#c5dff0');
  sheet.getRange(currentRow, 16).setBackground('#c5dff0');
  
  currentRow++;
  
  // Add borders around the entire table
  const tableRange = sheet.getRange(headerRow - 1, 1, currentRow - (headerRow - 1), numColumns);
  tableRange.setBorder(
    true, true, true, true, true, true,
    'black', SpreadsheetApp.BorderStyle.SOLID
  );
  
  // Add thicker borders around paired variance columns (same pattern as team table)
  const tableHeight = currentRow - (headerRow - 1);
  
  // ENTIRE PI pairs
  sheet.getRange(headerRow - 1, 4, tableHeight, 2).setBorder(true, true, true, true, false, false, 'black', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sheet.getRange(headerRow - 1, 6, tableHeight, 2).setBorder(true, true, true, true, false, false, 'black', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sheet.getRange(headerRow - 1, 8, tableHeight, 2).setBorder(true, true, true, true, false, false, 'black', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  
  // CODE FREEZE pairs
  sheet.getRange(headerRow - 1, 11, tableHeight, 2).setBorder(true, true, true, true, false, false, 'black', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sheet.getRange(headerRow - 1, 13, tableHeight, 2).setBorder(true, true, true, true, false, false, 'black', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sheet.getRange(headerRow - 1, 15, tableHeight, 2).setBorder(true, true, true, true, false, false, 'black', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  
  console.log(`Role capacity table created with ${sortedRoles.length} roles`);
  
  return currentRow;
}
function getRoleCapacityData(spreadsheet, piNumber) {
  const capacitySheetName = `PI${piNumber} - Capacity`;
  const capacitySheet = spreadsheet.getSheetByName(capacitySheetName);
  
  if (!capacitySheet) {
    console.warn(`Role capacity sheet "${capacitySheetName}" not found`);
    return { entirePI: {}, codeFreeze: {}, teamRoles: {} };
  }
  
  console.log(`\n=== Reading Role Capacity Data from "${capacitySheetName}" ===`);
  
  const dataRange = capacitySheet.getDataRange();
  const values = dataRange.getValues();
  const maxRows = values.length;
  const maxCols = values[0] ? values[0].length : 0;
  
  // Find the EMA Clinical column by scanning row 1 for "EMA Clinical"
  let clinicalCol = -1;
  for (let col = 0; col < maxCols; col++) {
    const headerVal = values[0][col];
    if (headerVal && headerVal.toString().trim().toLowerCase().includes('ema clinical')) {
      clinicalCol = col;
      break;
    }
  }
  
  if (clinicalCol === -1) {
    console.warn('EMA Clinical section not found in capacity sheet row 1');
    return { entirePI: {}, codeFreeze: {}, teamRoles: {} };
  }
  
  console.log(`Found EMA Clinical at column ${clinicalCol + 1} (${String.fromCharCode(65 + (clinicalCol % 26))})`);
  
  // The Total column is at offset +9 from the team column
  const totalColOffset = 9;
  
  // Aggregated role capacity across all teams
  const entirePIRoles = {};
  const codeFreezeRoles = {};
  const teamRoles = {};
  
  // Scan for team blocks - team name is in the clinical column, followed by structured data
  let row = 2; // Start after header row
  while (row < maxRows) {
    const cellValue = values[row][clinicalCol];
    
    // Skip empty rows and known non-team markers
    if (!cellValue || cellValue.toString().trim() === '' || cellValue.toString().trim() === '-') {
      row++;
      continue;
    }
    
    const cellStr = cellValue.toString().trim();
    
    // Skip known structural markers
    if (cellStr.toLowerCase().includes('allocation type') ||
        cellStr.toLowerCase().includes('base capacity') ||
        cellStr.toLowerCase().includes('klo') ||
        cellStr.toLowerCase().includes('quality') ||
        cellStr.toLowerCase().includes('tech') ||
        cellStr.toLowerCase().includes('product') ||
        cellStr.toLowerCase().includes('unplanned')) {
      row++;
      continue;
    }
    
    // Check if next row is "Allocation Type" - confirms this is a team name
    if (row + 1 < maxRows) {
      const nextRowVal = values[row + 1][clinicalCol];
      if (nextRowVal && nextRowVal.toString().trim() === 'Allocation Type') {
        const teamName = cellStr.toUpperCase();
        
        // Check if this is a clinical team we care about
        const isClinicalTeam = DANS_REPORT_CONFIG.clinicalTeams.some(
          t => t === teamName || t === cellStr.toUpperCase().replace(/\s+/g, ' ')
        );
        
        if (!isClinicalTeam) {
          // Not a clinical team, skip this block
          row += 25; // Jump past this team block
          continue;
        }
        
        console.log(`Processing clinical team: ${teamName} at row ${row + 1}`);
        teamRoles[teamName] = {};
        
        // Find "Base Capacity before FF" within the next 12 rows
        let beforeFFRow = -1;
        for (let offset = 8; offset <= 12; offset++) {
          if (row + offset < maxRows) {
            const checkVal = values[row + offset][clinicalCol];
            if (checkVal && checkVal.toString().toLowerCase().includes('before ff')) {
              beforeFFRow = row + offset;
              break;
            }
          }
        }
        
        // Find "Base Capacity after FF" within the next 20 rows
        let afterFFRow = -1;
        for (let offset = 16; offset <= 20; offset++) {
          if (row + offset < maxRows) {
            const checkVal = values[row + offset][clinicalCol];
            if (checkVal && checkVal.toString().toLowerCase().includes('after ff')) {
              afterFFRow = row + offset;
              break;
            }
          }
        }
        
        // Read Before FF roles (rows after beforeFFRow header)
        if (beforeFFRow > 0) {
          for (let rOffset = 1; rOffset <= 6; rOffset++) {
            const roleRow = beforeFFRow + rOffset;
            if (roleRow >= maxRows) break;
            
            const roleName = values[roleRow][clinicalCol];
            if (!roleName || roleName.toString().trim() === '' || roleName.toString().trim() === '-') continue;
            
            // Stop if we hit a subtotal row (no role name, just numbers)
            const roleStr = roleName.toString().trim();
            if (roleStr.toLowerCase().includes('base capacity')) break;
            
            const roleKey = roleStr.toUpperCase();
            const normalizedRole = ROLE_NORMALIZATION_DANS[roleKey] || roleKey;
            
            // Get the total value (at totalColOffset from clinicalCol)
            const totalVal = values[roleRow][clinicalCol + totalColOffset];
            let beforeFFTotal = 0;
            if (totalVal && totalVal !== '-' && totalVal !== '') {
              beforeFFTotal = parseFloat(totalVal) || 0;
            }
            
            if (!teamRoles[teamName][normalizedRole]) {
              teamRoles[teamName][normalizedRole] = { beforeFF: 0, afterFF: 0, entirePI: 0 };
            }
            teamRoles[teamName][normalizedRole].beforeFF = Math.ceil(beforeFFTotal);
          }
        }
        
        // Read After FF roles (rows after afterFFRow header)
        if (afterFFRow > 0) {
          for (let rOffset = 1; rOffset <= 6; rOffset++) {
            const roleRow = afterFFRow + rOffset;
            if (roleRow >= maxRows) break;
            
            const roleName = values[roleRow][clinicalCol];
            if (!roleName || roleName.toString().trim() === '' || roleName.toString().trim() === '-') continue;
            
            const roleStr = roleName.toString().trim();
            if (roleStr.toLowerCase().includes('base capacity')) break;
            
            const roleKey = roleStr.toUpperCase();
            const normalizedRole = ROLE_NORMALIZATION_DANS[roleKey] || roleKey;
            
            const totalVal = values[roleRow][clinicalCol + totalColOffset];
            let afterFFTotal = 0;
            if (totalVal && totalVal !== '-' && totalVal !== '') {
              afterFFTotal = parseFloat(totalVal) || 0;
            }
            
            if (!teamRoles[teamName][normalizedRole]) {
              teamRoles[teamName][normalizedRole] = { beforeFF: 0, afterFF: 0, entirePI: 0 };
            }
            teamRoles[teamName][normalizedRole].afterFF = Math.ceil(afterFFTotal);
          }
        }
        
        // Calculate Entire PI totals and aggregate across teams
        Object.keys(teamRoles[teamName]).forEach(role => {
          const rd = teamRoles[teamName][role];
          rd.entirePI = rd.beforeFF + rd.afterFF;
          
          if (!entirePIRoles[role]) entirePIRoles[role] = 0;
          if (!codeFreezeRoles[role]) codeFreezeRoles[role] = 0;
          
          entirePIRoles[role] += rd.entirePI;
          codeFreezeRoles[role] += rd.beforeFF; // Code Freeze = Before FF only
        });
        
        console.log(`  Roles found: ${Object.keys(teamRoles[teamName]).filter(r => teamRoles[teamName][r].entirePI > 0).join(', ')}`);
        
        // Move past this team block
        row += 25;
        continue;
      }
    }
    
    row++;
  }
  
  // Log summary
  console.log('\n=== Aggregated Role Capacity Summary ===');
  Object.keys(entirePIRoles).sort().forEach(role => {
    if (entirePIRoles[role] > 0 || codeFreezeRoles[role] > 0) {
      console.log(`  ${role}: Entire PI = ${entirePIRoles[role]}, Code Freeze = ${codeFreezeRoles[role]}`);
    }
  });
  console.log('=== End Role Capacity ===\n');
  
  return { entirePI: entirePIRoles, codeFreeze: codeFreezeRoles, teamRoles: teamRoles };
}

/**
 * Get capacity used (story points) by ROLE from the PI data sheet
 * Uses role detection (labels + title prefix) to assign story points to roles
 * 
 * @param {Sheet} piSheet - The PI data sheet
 * @return {Object} { ROLE: totalStoryPoints, ... }
 */
function getCapacityUsedByRole(piSheet) {
  if (!piSheet) {
    console.warn('PI sheet not provided - role capacity used will be 0');
    return {};
  }
  
  console.log(`\n=== Calculating Capacity Used by Role ===`);
  
  const dataRange = piSheet.getDataRange();
  const values = dataRange.getValues();
  
  if (values.length < 4) return {};
  
  const headers = values[3];
  const scrumTeamCol = headers.indexOf('Scrum Team');
  const storyPointsCol = headers.indexOf('Story Points');
  const issueTypeCol = headers.indexOf('Issue Type');
  const allocationCol = headers.indexOf('Allocation');
  const labelsCol = headers.indexOf('Labels');
  const summaryCol = headers.indexOf('Summary');
  
  if (scrumTeamCol === -1 || storyPointsCol === -1) {
    console.error('Required columns not found for role capacity used');
    return {};
  }
  
  const roleUsed = {};
  let matched = 0;
  let unmatched = 0;
  
  for (let i = 4; i < values.length; i++) {
    const row = values[i];
    const scrumTeam = row[scrumTeamCol];
    const storyPoints = parseFloat(row[storyPointsCol]) || 0;
    const issueType = issueTypeCol !== -1 ? row[issueTypeCol] : '';
    const allocation = allocationCol !== -1 ? row[allocationCol] : '';
    
    if (!scrumTeam || scrumTeam.toString().trim() === '') continue;
    if (issueType !== 'Story' && issueType !== 'Bug') continue;
    if (allocation !== 'Product - Feature' && allocation !== 'Product - Compliance') continue;
    if (storyPoints === 0) continue;
    
    // Only count clinical teams
    const cleanTeamName = scrumTeam.toString().trim().toUpperCase();
    if (!DANS_REPORT_CONFIG.clinicalTeams.includes(cleanTeamName)) continue;
    
    // Detect role
    const labels = labelsCol !== -1 ? (row[labelsCol] || '').toString().split(',').map(l => l.trim()) : [];
    const summary = summaryCol !== -1 ? row[summaryCol] : '';
    const detectedRole = detectRoleFromTicketDans(labels, summary) || 'Unassigned';
    
    if (!roleUsed[detectedRole]) roleUsed[detectedRole] = 0;
    roleUsed[detectedRole] += Math.ceil(storyPoints);
    
    if (detectedRole === 'Unassigned') unmatched++;
    else matched++;
  }
  
  console.log(`Role detection: ${matched} matched, ${unmatched} unassigned`);
  console.log('Role Used:', JSON.stringify(roleUsed));
  console.log('=== End Capacity Used by Role ===\n');
  
  return roleUsed;
}

/**
 * Get load data (Planned Load FP and Actual Load SPE) by ROLE from Epic-level data
 * 
 * For Epics, role is determined by looking at the Epic's own labels/title prefix.
 * If an Epic has no detectable role, its load is assigned to 'Unassigned'.
 * 
 * @param {Sheet} piSheet - The PI data sheet
 * @return {Object} { plannedLoad: { ROLE: value }, actualLoad: { ROLE: value } }
 */
function getLoadDataByRole(piSheet) {
  if (!piSheet) {
    return { plannedLoad: {}, actualLoad: {} };
  }
  
  console.log(`\n=== Calculating Load Data by Role ===`);
  
  const dataRange = piSheet.getDataRange();
  const values = dataRange.getValues();
  
  if (values.length < 4) return { plannedLoad: {}, actualLoad: {} };
  
  const headers = values[3];
  const scrumTeamCol = headers.indexOf('Scrum Team');
  const issueTypeCol = headers.indexOf('Issue Type');
  const allocationCol = headers.indexOf('Allocation');
  const featurePointsCol = headers.indexOf('Feature Points');
  const storyPointEstimateCol = headers.indexOf('Story Point Estimate');
  const labelsCol = headers.indexOf('Labels');
  const summaryCol = headers.indexOf('Summary');
  
  if (scrumTeamCol === -1 || issueTypeCol === -1 || allocationCol === -1) {
    return { plannedLoad: {}, actualLoad: {} };
  }
  
  const plannedLoad = {};
  const actualLoad = {};
  
  for (let i = 4; i < values.length; i++) {
    const row = values[i];
    const scrumTeam = row[scrumTeamCol];
    const issueType = row[issueTypeCol];
    const allocation = row[allocationCol];
    
    if (!scrumTeam || scrumTeam.toString().trim() === '') continue;
    if (issueType !== 'Epic') continue;
    if (allocation !== 'Product - Feature' && allocation !== 'Product - Compliance') continue;
    
    const cleanTeamName = scrumTeam.toString().trim().toUpperCase();
    if (!DANS_REPORT_CONFIG.clinicalTeams.includes(cleanTeamName)) continue;
    
    // Detect role from Epic labels/title
    const labels = labelsCol !== -1 ? (row[labelsCol] || '').toString().split(',').map(l => l.trim()) : [];
    const summary = summaryCol !== -1 ? row[summaryCol] : '';
    const detectedRole = detectRoleFromTicketDans(labels, summary) || 'Unassigned';
    
    const featurePoints = featurePointsCol !== -1 ? parseFloat(row[featurePointsCol]) || 0 : 0;
    const storyPointEstimate = storyPointEstimateCol !== -1 ? parseFloat(row[storyPointEstimateCol]) || 0 : 0;
    
    if (!plannedLoad[detectedRole]) plannedLoad[detectedRole] = 0;
    if (!actualLoad[detectedRole]) actualLoad[detectedRole] = 0;
    
    if (featurePoints > 0) {
      plannedLoad[detectedRole] += Math.ceil(featurePoints) * 10;
    }
    if (storyPointEstimate > 0) {
      actualLoad[detectedRole] += Math.ceil(storyPointEstimate);
    }
  }
  
  console.log('Planned Load by Role (FP×10):', JSON.stringify(plannedLoad));
  console.log('Actual Load by Role (SPE):', JSON.stringify(actualLoad));
  console.log('=== End Load Data by Role ===\n');
  
  return { plannedLoad, actualLoad };
}

/**
 * Create the role-based capacity utilization table
 * Same 16-column structure as the team-based table but with roles as rows
 * 
 * @param {Sheet} sheet - The report sheet
 * @param {number} currentRow - Starting row for this table
 * @param {Object} roleCapacity - From getRoleCapacityData()
 * @param {Object} roleUsed - From getCapacityUsedByRole()
 * @param {Object} roleLoadData - From getLoadDataByRole()
 * @return {number} Next available row after the table
 */
function createRoleCapacityTable(sheet, currentRow, roleCapacity, roleUsed, roleLoadData) {
  const numColumns = 16;
  
  // Section title
  sheet.getRange(currentRow, 1).setValue('CLINICAL CAPACITY UTILIZATION BY ROLE - ENTIRE PI vs CODE FREEZE');
  sheet.getRange(currentRow, 1)
    .setFontSize(14)
    .setFontWeight('bold')
    .setFontFamily('Comfortaa')
    .setBackground('#E1D5E7')
    .setFontColor('black');
  sheet.getRange(currentRow, 1, 1, numColumns).merge();
  currentRow++;
  
  // Category headers
  sheet.getRange(currentRow, 1).setValue('');
  sheet.getRange(currentRow, 2).setValue('');
  sheet.getRange(currentRow, 3, 1, 7).merge();
  sheet.getRange(currentRow, 3).setValue('ENTIRE PI');
  sheet.getRange(currentRow, 10, 1, 7).merge();
  sheet.getRange(currentRow, 10).setValue('CODE FREEZE');
  
  sheet.getRange(currentRow, 1, 1, numColumns)
    .setFontWeight('bold')
    .setBackground('#D5A6E0')
    .setFontColor('black')
    .setFontSize(10)
    .setFontFamily('Comfortaa')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  currentRow++;
  
  // Column headers (same as team table but "Role" instead of "Scrum Team")
  const headers = [
    'Role',
    'Baseline Capacity',
    'Product Capacity',
    'Capacity Used (LOE)',
    'Capacity Remaining',
    'Planned Load (FP)',
    'Planned Remaining',
    'Actual Load (SPE)',
    'Actual Remaining',
    'Product Capacity',
    'Capacity Used (LOE)',
    'Capacity Remaining',
    'Planned Load (FP)',
    'Planned Remaining',
    'Actual Load (SPE)',
    'Actual Remaining'
  ];
  
  sheet.getRange(currentRow, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(currentRow, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#9b7bb8')
    .setFontColor('white')
    .setFontSize(9)
    .setFontFamily('Comfortaa')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true);
  
  // Add notes to role table headers
  sheet.getRange(currentRow, 2).setNote('Sum of Before FF + After FF capacity for this role across all clinical teams');
  sheet.getRange(currentRow, 3).setNote('ENTIRE PI: Before FF + After FF capacity for this role\nAggregated across all clinical teams');
  sheet.getRange(currentRow, 4).setNote('Story points from Stories/Bugs assigned to this role\n(detected from labels or title prefix)');
  sheet.getRange(currentRow, 10).setNote('CODE FREEZE: Before FF capacity only (iterations 1-5)\nLIGHT BLUE = Different data source');
  
  const headerRow = currentRow;
  currentRow++;
  
  // Determine which roles to display
  // Collect all roles from capacity and usage, sorted with Unassigned last
  const allRoles = new Set();
  Object.keys(roleCapacity.entirePI).forEach(r => { if (roleCapacity.entirePI[r] > 0) allRoles.add(r); });
  Object.keys(roleCapacity.codeFreeze).forEach(r => { if (roleCapacity.codeFreeze[r] > 0) allRoles.add(r); });
  Object.keys(roleUsed).forEach(r => allRoles.add(r));
  Object.keys(roleLoadData.plannedLoad).forEach(r => allRoles.add(r));
  Object.keys(roleLoadData.actualLoad).forEach(r => allRoles.add(r));
  
  // Remove Unassigned from main list (will be added at end)
  allRoles.delete('Unassigned');
  
  // Sort roles alphabetically
  const sortedRoles = Array.from(allRoles).sort();
  
  // Add Unassigned at the end if it has any data
  const hasUnassigned = (roleUsed['Unassigned'] || 0) > 0 || 
                        (roleLoadData.plannedLoad['Unassigned'] || 0) > 0 ||
                        (roleLoadData.actualLoad['Unassigned'] || 0) > 0;
  if (hasUnassigned) {
    sortedRoles.push('Unassigned');
  }
  
  const dataStartRow = currentRow;
  
  // Write data rows for each role
  for (const role of sortedRoles) {
    const isUnassigned = role === 'Unassigned';
    
    // Capacity values
    const entirePICap = roleCapacity.entirePI[role] || 0;
    const codeFreezeCap = roleCapacity.codeFreeze[role] || 0;
    const capacityUsed = roleUsed[role] || 0;
    
    // Baseline Capacity = Entire PI capacity (same concept as team table)
    const sprintCapacity = entirePICap;
    
    // Remaining
    const remainingEntirePI = entirePICap - capacityUsed;
    const remainingCodeFreeze = codeFreezeCap - capacityUsed;
    
    // Load data
    const plannedLoad = roleLoadData.plannedLoad[role] || 0;
    const actualLoad = roleLoadData.actualLoad[role] || 0;
    
    // Planned/Actual Remaining
    const plannedRemainingEntirePI = entirePICap - plannedLoad;
    const actualRemainingEntirePI = entirePICap - actualLoad;
    const plannedRemainingCodeFreeze = codeFreezeCap - plannedLoad;
    const actualRemainingCodeFreeze = codeFreezeCap - actualLoad;
    
    // Write row
    sheet.getRange(currentRow, 1).setValue(role);
    sheet.getRange(currentRow, 2).setValue(sprintCapacity);
    
    // ENTIRE PI section
    sheet.getRange(currentRow, 3).setValue(entirePICap);
    sheet.getRange(currentRow, 4).setValue(capacityUsed);
    sheet.getRange(currentRow, 5).setValue(remainingEntirePI);
    sheet.getRange(currentRow, 6).setValue(plannedLoad);
    sheet.getRange(currentRow, 7).setValue(plannedRemainingEntirePI);
    sheet.getRange(currentRow, 8).setValue(actualLoad);
    sheet.getRange(currentRow, 9).setValue(actualRemainingEntirePI);
    
    // CODE FREEZE section
    sheet.getRange(currentRow, 10).setValue(codeFreezeCap);
    sheet.getRange(currentRow, 11).setValue(capacityUsed);
    sheet.getRange(currentRow, 12).setValue(remainingCodeFreeze);
    sheet.getRange(currentRow, 13).setValue(plannedLoad);
    sheet.getRange(currentRow, 14).setValue(plannedRemainingCodeFreeze);
    sheet.getRange(currentRow, 15).setValue(actualLoad);
    sheet.getRange(currentRow, 16).setValue(actualRemainingCodeFreeze);
    
    // Format row
    sheet.getRange(currentRow, 1, 1, numColumns)
      .setFontSize(9)
      .setFontFamily('Comfortaa')
      .setVerticalAlignment('middle');
    
    sheet.getRange(currentRow, 2, 1, 15)
      .setHorizontalAlignment('center')
      .setNumberFormat('#,##0');
    
    // Light grey for Baseline Capacity
    sheet.getRange(currentRow, 2).setBackground('#f5f5f5');
    
    // Light blue for CODE FREEZE columns with different data source
    sheet.getRange(currentRow, 10).setBackground('#e6f2ff');
    sheet.getRange(currentRow, 12).setBackground('#e6f2ff');
    sheet.getRange(currentRow, 14).setBackground('#e6f2ff');
    sheet.getRange(currentRow, 16).setBackground('#e6f2ff');
    
    // Yellow highlight for Unassigned
    if (isUnassigned) {
      sheet.getRange(currentRow, 1).setBackground('#fff3cd');
      if (capacityUsed > 0) {
        sheet.getRange(currentRow, 4).setBackground('#fff3cd');
        sheet.getRange(currentRow, 11).setBackground('#fff3cd');
      }
    }
    
    // Conditional formatting for remaining columns - ENTIRE PI
    const remainEntirePICell = sheet.getRange(currentRow, 5);
    if (remainingEntirePI < 0) {
      remainEntirePICell.setBackground('#ffcccc').setFontColor('#cc0000').setFontWeight('bold');
    } else {
      remainEntirePICell.setBackground('#d4edda').setFontColor('#155724').setFontWeight('bold');
    }
    
    const plannedRemEntirePICell = sheet.getRange(currentRow, 7);
    if (plannedRemainingEntirePI < 0) {
      plannedRemEntirePICell.setBackground('#ffcccc').setFontColor('#cc0000').setFontWeight('bold');
    } else {
      plannedRemEntirePICell.setBackground('#d4edda').setFontColor('#155724').setFontWeight('bold');
    }
    
    const actualRemEntirePICell = sheet.getRange(currentRow, 9);
    if (actualRemainingEntirePI < 0) {
      actualRemEntirePICell.setBackground('#ffcccc').setFontColor('#cc0000').setFontWeight('bold');
    } else {
      actualRemEntirePICell.setBackground('#d4edda').setFontColor('#155724').setFontWeight('bold');
    }
    
    // Conditional formatting for remaining columns - CODE FREEZE (blended colors)
    const remainCFCell = sheet.getRange(currentRow, 12);
    if (remainingCodeFreeze < 0) {
      remainCFCell.setBackground('#ffccdd').setFontColor('#cc0000').setFontWeight('bold');
    } else {
      remainCFCell.setBackground('#ccf2e6').setFontColor('#155724').setFontWeight('bold');
    }
    
    const plannedRemCFCell = sheet.getRange(currentRow, 14);
    if (plannedRemainingCodeFreeze < 0) {
      plannedRemCFCell.setBackground('#ffccdd').setFontColor('#cc0000').setFontWeight('bold');
    } else {
      plannedRemCFCell.setBackground('#ccf2e6').setFontColor('#155724').setFontWeight('bold');
    }
    
    const actualRemCFCell = sheet.getRange(currentRow, 16);
    if (actualRemainingCodeFreeze < 0) {
      actualRemCFCell.setBackground('#ffccdd').setFontColor('#cc0000').setFontWeight('bold');
    } else {
      actualRemCFCell.setBackground('#ccf2e6').setFontColor('#155724').setFontWeight('bold');
    }
    
    currentRow++;
  }
  
  const dataEndRow = currentRow - 1;
  
  // Add TOTAL row
  sheet.getRange(currentRow, 1).setValue('TOTAL');
  
  for (let col = 2; col <= 16; col++) {
    const colLetter = String.fromCharCode(64 + col); // B=66, C=67, etc.
    sheet.getRange(currentRow, col).setFormula(`=SUM(${colLetter}${dataStartRow}:${colLetter}${dataEndRow})`);
  }
  
  // Format totals row
  sheet.getRange(currentRow, 1, 1, numColumns)
    .setFontWeight('bold')
    .setFontFamily('Comfortaa')
    .setBackground('#f0f0f0');
  
  sheet.getRange(currentRow, 2, 1, 15)
    .setHorizontalAlignment('center')
    .setNumberFormat('#,##0');
  
  sheet.getRange(currentRow, 2).setBackground('#d9d9d9');
  sheet.getRange(currentRow, 10).setBackground('#c5dff0');
  sheet.getRange(currentRow, 12).setBackground('#c5dff0');
  sheet.getRange(currentRow, 14).setBackground('#c5dff0');
  sheet.getRange(currentRow, 16).setBackground('#c5dff0');
  
  currentRow++;
  
  // Add borders around the entire table
  const tableRange = sheet.getRange(headerRow - 1, 1, currentRow - (headerRow - 1), numColumns);
  tableRange.setBorder(
    true, true, true, true, true, true,
    'black', SpreadsheetApp.BorderStyle.SOLID
  );
  
  // Add thicker borders around paired variance columns (same pattern as team table)
  const tableHeight = currentRow - (headerRow - 1);
  
  // ENTIRE PI pairs
  sheet.getRange(headerRow - 1, 4, tableHeight, 2).setBorder(true, true, true, true, false, false, 'black', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sheet.getRange(headerRow - 1, 6, tableHeight, 2).setBorder(true, true, true, true, false, false, 'black', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sheet.getRange(headerRow - 1, 8, tableHeight, 2).setBorder(true, true, true, true, false, false, 'black', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  
  // CODE FREEZE pairs
  sheet.getRange(headerRow - 1, 11, tableHeight, 2).setBorder(true, true, true, true, false, false, 'black', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sheet.getRange(headerRow - 1, 13, tableHeight, 2).setBorder(true, true, true, true, false, false, 'black', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sheet.getRange(headerRow - 1, 15, tableHeight, 2).setBorder(true, true, true, true, false, false, 'black', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  
  console.log(`Role capacity table created with ${sortedRoles.length} roles`);
  
  return currentRow;
}

// ===== MAIN REPORT GENERATION =====
/**
 * Create the capacity utilization report with all tables
 * @param {Sheet} sheet - The report sheet
 * @param {Object} capacityDataEntirePI - Capacity data for Entire PI (rows 36-47)
 * @param {Object} capacityDataCodeFreeze - Capacity data for Code Freeze (rows 4-15)
 * @param {Object} piData - PI data for capacity used
 * @param {string} programIncrement - The PI name (e.g., "PI 13")
 * @param {Sheet} piSheet - The PI data sheet (for load calculations and blank fix version table)
 * @param {string} piNumber - The PI number (e.g., "13")
 */
function createCapacityUtilizationReport(sheet, capacityDataEntirePI, capacityDataCodeFreeze, piData, programIncrement, piSheet, piNumber) {
  let currentRow = 1;

  // Add title
  sheet.getRange(currentRow, 1).setValue(`DAN'S REPORT - CLINICAL CAPACITY UTILIZATION (${programIncrement})`);
  sheet.getRange(currentRow, 1)
    .setFontSize(16)
    .setFontWeight('bold')
    .setFontFamily('Comfortaa')
    .setBackground('#4A235A')
    .setFontColor('white');
  sheet.getRange(currentRow, 1, 1, 16).merge();
  currentRow++;

  // Add Report Generated timestamp
  const reportTimestamp = new Date();
  const formattedReportTime = Utilities.formatDate(reportTimestamp, Session.getScriptTimeZone(), 'MMM dd, yyyy HH:mm:ss');
  sheet.getRange(currentRow, 1).setValue(`Report Generated: ${formattedReportTime}`);
  sheet.getRange(currentRow, 1)
    .setFontSize(10)
    .setFontStyle('italic')
    .setFontFamily('Comfortaa');
  sheet.getRange(currentRow, 1, 1, 16).merge();
  currentRow++;

  // Add PI Data Last Refreshed timestamp (from PI sheet)
  let piDataTimestamp = 'Unknown';
  if (piSheet) {
    try {
      // Read timestamp from PI sheet (Row 2, Column B)
      const piTimestampValue = piSheet.getRange(2, 2).getValue();
      if (piTimestampValue) {
        // If it's already a formatted string, use it directly
        if (typeof piTimestampValue === 'string') {
          piDataTimestamp = piTimestampValue;
        } else if (piTimestampValue instanceof Date) {
          // If it's a Date object, format it
          piDataTimestamp = Utilities.formatDate(piTimestampValue, Session.getScriptTimeZone(), 'MMM dd, yyyy HH:mm:ss');
        }
      }
    } catch (error) {
      console.error('Error reading PI sheet timestamp:', error);
      piDataTimestamp = 'Error reading timestamp';
    }
  }

  sheet.getRange(currentRow, 1).setValue(`PI Data Last Refreshed: ${piDataTimestamp}`);
  sheet.getRange(currentRow, 1)
    .setFontSize(10)
    .setFontStyle('italic')
    .setFontFamily('Comfortaa')
    .setFontColor('#666666'); // Slightly darker to differentiate from report timestamp
  sheet.getRange(currentRow, 1, 1, 16).merge();
  currentRow++;

  // Add blank row
  currentRow++;

  // Get load data for the table
  const loadData = getLoadData(SpreadsheetApp.getActiveSpreadsheet(), piSheet, piNumber);

  // Create combined table with Entire PI and Code Freeze side-by-side
  currentRow = createCombinedCapacityTable(sheet, currentRow, capacityDataEntirePI, capacityDataCodeFreeze, piData, loadData);

  // Add spacing before role-based table
  currentRow += 2;

  // Create role-based capacity table (same 16-column structure, roles as rows)
  const roleCapacity = getRoleCapacityData(SpreadsheetApp.getActiveSpreadsheet(), piNumber);
  const roleUsed = getCapacityUsedByRole(piSheet);
  const roleLoadData = getLoadDataByRole(piSheet);

  // Only create role table if we have role data
  if (Object.keys(roleCapacity.entirePI).length > 0 || Object.keys(roleUsed).length > 0) {
    currentRow = createRoleCapacityTable(sheet, currentRow, roleCapacity, roleUsed, roleLoadData);
  } else {
    console.log('No role capacity data found - skipping role-based table');
  }

  // Add spacing before blank fix version table
  currentRow += 2;

  // Create blank fix version table
  currentRow = createBlankFixVersionTable(sheet, currentRow, piSheet);

  // Add spacing before role breakdown section
  currentRow += 2;

  // Create Role Breakdown section (if roles exist in capacity sheet)
  currentRow = createDansReportRoleBreakdown(sheet, currentRow, SpreadsheetApp.getActiveSpreadsheet(), piSheet);
}
/**
 * Format the overall sheet appearance
 */
function formatDansReportSheet(sheet) {
  // Set column widths for the combined 16-column table
  sheet.setColumnWidth(1, 130);  // Scrum Team
  sheet.setColumnWidth(2, 90);   // Baseline Capacity
  
  // ENTIRE PI section (columns 3-9)
  sheet.setColumnWidth(3, 85);   // Product Capacity
  sheet.setColumnWidth(4, 85);   // Capacity Used
  sheet.setColumnWidth(5, 85);   // Capacity Remaining
  sheet.setColumnWidth(6, 95);   // Planned Load (FP)
  sheet.setColumnWidth(7, 95);   // Planned Remaining
  sheet.setColumnWidth(8, 90);   // Actual Load (SPE)
  sheet.setColumnWidth(9, 90);   // Actual Remaining
  
  // CODE FREEZE section (columns 10-16)
  sheet.setColumnWidth(10, 85);  // Product Capacity
  sheet.setColumnWidth(11, 85);  // Capacity Used
  sheet.setColumnWidth(12, 85);  // Capacity Remaining
  sheet.setColumnWidth(13, 95);  // Planned Load (FP)
  sheet.setColumnWidth(14, 95);  // Planned Remaining
  sheet.setColumnWidth(15, 90);  // Actual Load (SPE)
  sheet.setColumnWidth(16, 90);  // Actual Remaining
  
  // Note: Borders and text wrapping are applied within the table creation function
  
  // Freeze the header rows (first 4 rows - title, report timestamp, PI data timestamp, blank)
  sheet.setFrozenRows(4);
}
