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
 * - Planned Load (FP): At Epic level, sum Feature Points x 10 where:
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
  
  // Capacity sheet name pattern - will try PI-specific name first
  // Example: "PI14 - Capacity", then fallback to "Clinical: Capacity Planning"
  capacitySheetPattern: 'PI{PI_NUMBER} - Capacity',
  capacitySheetFallback: 'Clinical: Capacity Planning',
  
  // JIRA base URL for creating hyperlinks
  jiraBaseUrl: 'https://modmedrnd.atlassian.net',
  
  // Clinical scrum teams to include (in display order, UPPERCASE for matching)
  // Updated to match new consolidated capacity format
  clinicalTeams: [
    'ALCHEMIST',
    'AVENGERS',
    'EXPLORERS',
    'EYEFINITY',
    'MANDALORE',
    'ORDERNAUTS',
    'PAINKILLERS',
    'ARTIFICIALLY INTELLIGENT',
    'PATIENCE',
    // Legacy teams (may still appear in JIRA data)
    'EMBRYONICS',
    'VESTIES',
    'SPICE RUNNERS',
    'PAIN KILLERS'
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
    
    // Remove existing charts (clear() doesn't remove them)
    const existingCharts = reportSheet.getCharts();
    existingCharts.forEach(chart => reportSheet.removeChart(chart));
    
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

/**
 * Generate Dan's Report programmatically (no prompts)
 * Used by the unified analysis when generating a new report
 * @param {Spreadsheet} spreadsheet - Target spreadsheet
 * @param {number|string} piNumber - PI number
 * @returns {boolean} Success status
 */
function generateDansReportForSpreadsheet(spreadsheet, piNumber, sourceSpreadsheet) {
  try {
    console.log(`Generating Dan's Report for PI ${piNumber} (programmatic call)`);
    
    const programIncrement = `PI ${piNumber}`;
    
    // Use source spreadsheet for capacity data if provided, otherwise use target
    const capacitySpreadsheet = sourceSpreadsheet || spreadsheet;
    console.log(`Using ${sourceSpreadsheet ? 'source' : 'target'} spreadsheet for capacity data`);
    
    // Check if PI data sheet exists in target spreadsheet
    const piSheet = spreadsheet.getSheetByName(programIncrement);
    if (!piSheet) {
      console.error(`Sheet "${programIncrement}" not found for Dan's Report`);
      return false;
    }
    
    // Create or get the report sheet in target spreadsheet
    let reportSheet = spreadsheet.getSheetByName(DANS_REPORT_CONFIG.reportSheetName);
    if (!reportSheet) {
      reportSheet = spreadsheet.insertSheet(DANS_REPORT_CONFIG.reportSheetName);
    }
    
    // Remove existing charts (clear() doesn't remove them)
    const existingCharts = reportSheet.getCharts();
    existingCharts.forEach(chart => reportSheet.removeChart(chart));
    
    // Clear existing content
    reportSheet.clear();
    
    // Get capacity data from SOURCE spreadsheet (or target if source not provided)
    console.log('Reading Entire PI capacity data...');
    const capacityDataEntirePI = getClinicalCapacityData(capacitySpreadsheet, 36, 47, piNumber);
    
    console.log('Reading Code Freeze capacity data...');
    const capacityDataCodeFreeze = getClinicalCapacityData(capacitySpreadsheet, 4, 15, piNumber);
    
    // Get PI data for capacity used calculations from TARGET spreadsheet
    const piData = getPIDataForCapacityUsed(spreadsheet, piSheet);
    
    // Generate the report with both tables and blank fix version table
    createCapacityUtilizationReport(reportSheet, capacityDataEntirePI, capacityDataCodeFreeze, piData, programIncrement, piSheet, piNumber);
    
    // Format the sheet
    formatDansReportSheet(reportSheet);
    
    console.log(`Dan's Report generated successfully for ${programIncrement}`);
    return true;
    
  } catch (error) {
    console.error('Error generating Dan\'s Report (programmatic):', error);
    return false;
  }
}

// ===== DATA COLLECTION FUNCTIONS =====

/**
 * Find the capacity sheet - tries consolidated format first, then legacy
 * @param {Spreadsheet} spreadsheet - The spreadsheet object
 * @param {number} piNumber - Optional PI number for dynamic sheet name
 * @returns {Object} { sheet, isConsolidated } or throws error if not found
 */
function findDansReportCapacitySheet(spreadsheet, piNumber = null) {
  // Try consolidated format first (PI# - Capacity)
  if (typeof findCapacityPlanningSheet === 'function') {
    const consolidatedSheet = findCapacityPlanningSheet(spreadsheet, piNumber);
    if (consolidatedSheet) {
      return { sheet: consolidatedSheet, isConsolidated: true };
    }
  }
  
  // Try PI-specific sheet name
  if (piNumber) {
    const piSheetName = DANS_REPORT_CONFIG.capacitySheetPattern.replace('{PI_NUMBER}', piNumber);
    let sheet = spreadsheet.getSheetByName(piSheetName);
    if (sheet) {
      console.log(`Found capacity sheet: "${piSheetName}"`);
      // Check if it's consolidated format (Row 1 has value stream names)
      const row1 = sheet.getRange(1, 1, 1, 50).getValues()[0];
      const hasMultipleVS = row1.filter(v => v && v.toString().trim()).length > 1;
      return { sheet: sheet, isConsolidated: hasMultipleVS };
    }
  }
  
  // Fallback to legacy sheet name
  const legacySheet = spreadsheet.getSheetByName(DANS_REPORT_CONFIG.capacitySheetFallback);
  if (legacySheet) {
    console.log(`Using legacy capacity sheet: "${DANS_REPORT_CONFIG.capacitySheetFallback}"`);
    return { sheet: legacySheet, isConsolidated: false };
  }
  
  throw new Error('No capacity sheet found. Tried consolidated and legacy formats.');
}

/**
 * Get capacity data from the consolidated format for clinical teams
 * @param {Sheet} capacitySheet - The consolidated capacity sheet
 * @param {boolean} beforeFF - True for before Feature Freeze, false for after
 * @returns {Object} Capacity data by team
 */
function getClinicalCapacityDataConsolidated(capacitySheet, beforeFF = true) {
  const capacityData = {};
  const config = DANS_REPORT_CONFIG;
  
  console.log(`Reading consolidated capacity data (beforeFF: ${beforeFF})`);
  
  // Find EMA Clinical column in row 1
  const row1 = capacitySheet.getRange(1, 1, 1, 100).getValues()[0];
  let clinicalCol = -1;
  
  for (let col = 0; col < row1.length; col++) {
    const val = row1[col];
    if (val && val.toString().trim().toUpperCase().includes('EMA CLINICAL')) {
      clinicalCol = col + 1; // Convert to 1-indexed
      console.log(`Found EMA Clinical at column ${clinicalCol}`);
      break;
    }
  }
  
  if (clinicalCol === -1) {
    console.log('EMA Clinical column not found in consolidated sheet');
    return capacityData;
  }
  
  // Scan for teams under EMA Clinical column
  const teamBlockHeight = 25;
  let row = 3; // First team row
  const maxRows = capacitySheet.getLastRow();
  
  while (row <= maxRows) {
    const teamName = capacitySheet.getRange(row, clinicalCol).getValue();
    
    if (teamName && teamName.toString().trim()) {
      const cleanTeamName = teamName.toString().trim().toUpperCase();
      
      // Skip non-team rows
      if (!cleanTeamName.includes('ALLOCATION') && 
          !cleanTeamName.includes('BASE CAPACITY') &&
          !cleanTeamName.includes('MOBILE') &&
          !cleanTeamName.includes('W-DEV') &&
          !cleanTeamName.includes('QA')) {
        
        // Read capacity data from this team block
        // Total column is at offset +9 from team column
        const totalCol = clinicalCol + 9;
        
        // Base Capacity before FF total is at row offset +16
        // Base Capacity after FF total is at row offset +24
        const capacityRowOffset = beforeFF ? 16 : 24;
        
        // Product capacity = Product-Feature + Product-Compliance (rows +5 and +6)
        const productFeature = parseFloat(capacitySheet.getRange(row + 5, totalCol).getValue()) || 0;
        const productCompliance = parseFloat(capacitySheet.getRange(row + 6, totalCol).getValue()) || 0;
        
        // Get baseline from Base Capacity section
        let baseline = 0;
        const baseCapVal = capacitySheet.getRange(row + capacityRowOffset, totalCol).getValue();
        if (baseCapVal && baseCapVal !== '-') {
          baseline = Math.ceil(parseFloat(baseCapVal) || 0);
        }
        
        capacityData[cleanTeamName] = {
          baseline: baseline,
          productCapacity: Math.ceil(productFeature + productCompliance),
          productL: Math.ceil(productFeature),
          productM: Math.ceil(productCompliance)
        };
        
        console.log(`  Team: ${cleanTeamName}, Baseline: ${baseline}, Product: ${capacityData[cleanTeamName].productCapacity}`);
      }
    }
    
    row += teamBlockHeight;
  }
  
  return capacityData;
}

/**
 * Get capacity data from the Clinical: Capacity Planning sheet
 * Returns object with baseline and product capacity for each team
 * @param {Spreadsheet} spreadsheet - The spreadsheet object
 * @param {number} startRow - Starting row number (e.g., 36 for Entire PI, 4 for Code Freeze)
 * @param {number} endRow - Ending row number (e.g., 47 for Entire PI, 15 for Code Freeze)
 * @param {number} piNumber - Optional PI number for finding the right capacity sheet
 */
function getClinicalCapacityData(spreadsheet, startRow, endRow, piNumber = null) {
  // Try to find capacity sheet (consolidated or legacy)
  const { sheet: capacitySheet, isConsolidated } = findDansReportCapacitySheet(spreadsheet, piNumber);
  
  // If consolidated format, use new reader
  if (isConsolidated) {
    console.log('Using consolidated capacity format for clinical teams');
    // Determine if we're looking for before or after Feature Freeze
    const beforeFF = startRow > 20; // Entire PI starts at row 36
    return getClinicalCapacityDataConsolidated(capacitySheet, beforeFF);
  }
  
  // Legacy format handling
  const config = DANS_REPORT_CONFIG;
  const capacityData = {};
  
  console.log(`Reading legacy capacity data from rows ${startRow} to ${endRow}`);
  
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

// ===== ROLE BREAKDOWN FUNCTIONS FOR DAN'S REPORT =====

/**
 * Role name normalization mapping (same as scrumTeamSummary.js)
 */
const ROLE_NORMALIZATION_DANS = {
  'QA': 'QA', 'AQA': 'QA', 'QUALITY': 'QA',
  'W-DEV': 'W-DEV', 'WDEV': 'W-DEV', 'WEB': 'W-DEV', 'WEB-DEV': 'W-DEV', 'WEBDEV': 'W-DEV',
  'M-DEV': 'M-DEV', 'MDEV': 'M-DEV', 'MOBILE': 'M-DEV', 'MOBILE-DEV': 'M-DEV', 
  'M-ANDROID': 'M-DEV', 'M-IOS': 'M-DEV',
  'BE': 'BE', 'BACKEND': 'BE', 'BACK-END': 'BE',
  'FE': 'FE', 'FRONTEND': 'FE', 'FRONT-END': 'FE'
};

/**
 * Detect role from a ticket's labels or summary
 * @param {Array} labels - Array of labels
 * @param {string} summary - Ticket summary/title
 * @returns {string|null} Normalized role name or null
 */
function detectRoleFromTicketDans(labels, summary) {
  const labelArray = labels || [];
  const summaryStr = summary || '';
  
  const rolePatterns = [
    { pattern: /\bQA\b/i, role: 'QA' },
    { pattern: /\bAQA\b/i, role: 'QA' },
    { pattern: /\bBE\b/i, role: 'BE' },
    { pattern: /\bBACKEND\b/i, role: 'BE' },
    { pattern: /\bFE\b/i, role: 'FE' },
    { pattern: /\bFRONTEND\b/i, role: 'FE' },
    { pattern: /\bW-?DEV\b/i, role: 'W-DEV' },
    { pattern: /\bWEB\b/i, role: 'W-DEV' },
    { pattern: /\bM-?DEV\b/i, role: 'M-DEV' },
    { pattern: /\bMOBILE\b/i, role: 'M-DEV' },
    { pattern: /\bM-?ANDROID\b/i, role: 'M-DEV' },
    { pattern: /\bM-?IOS\b/i, role: 'M-DEV' }
  ];
  
  // Check labels first
  for (const label of labelArray) {
    const labelUpper = label.toString().toUpperCase().trim();
    if (ROLE_NORMALIZATION_DANS[labelUpper]) {
      return ROLE_NORMALIZATION_DANS[labelUpper];
    }
    for (const rp of rolePatterns) {
      if (rp.pattern.test(label)) return rp.role;
    }
  }
  
  // Check title prefix patterns
  const prefixPatterns = [
    /^\s*\[([A-Z\-]+)\]/i,
    /^\s*\(([A-Z\-]+)\)/i,
    /^\s*([A-Z\-]+)\s*:/i,
    /^\s*([A-Z\-]+)\s*-\s/i,
    /^\s*([A-Z\-]{2,6})\s+/i
  ];
  
  for (const pattern of prefixPatterns) {
    const match = summaryStr.match(pattern);
    if (match) {
      const extracted = match[1].toUpperCase().trim();
      if (ROLE_NORMALIZATION_DANS[extracted]) return ROLE_NORMALIZATION_DANS[extracted];
      for (const rp of rolePatterns) {
        if (rp.pattern.test(extracted)) return rp.role;
      }
    }
  }
  
  // Check for bracketed roles anywhere
  const bracketPatterns = [/\[([A-Z\-]+)\]/i, /\(([A-Z\-]+)\)/i];
  for (const pattern of bracketPatterns) {
    const match = summaryStr.match(pattern);
    if (match) {
      const extracted = match[1].toUpperCase().trim();
      if (ROLE_NORMALIZATION_DANS[extracted]) return ROLE_NORMALIZATION_DANS[extracted];
    }
  }
  
  return null;
}

/**
 * Get roles for a clinical team from the consolidated capacity sheet
 * @param {Spreadsheet} spreadsheet - The spreadsheet object
 * @param {string} teamName - Team name (e.g., "ALCHEMIST")
 * @returns {Object|null} { roles: { roleName: { beforeFF, afterFF, total } } }
 */
function getClinicalTeamRolesFromCapacity(spreadsheet, teamName) {
  try {
    // Try consolidated capacity sheet first
    let capacitySheet = null;
    
    if (typeof findCapacityPlanningSheet === 'function') {
      capacitySheet = findCapacityPlanningSheet(spreadsheet);
    }
    
    if (!capacitySheet) {
      const sheetNames = ['Capacity Planning', 'PI14 - Capacity', 'PI15 - Capacity'];
      for (const name of sheetNames) {
        capacitySheet = spreadsheet.getSheetByName(name);
        if (capacitySheet) break;
      }
      if (!capacitySheet) {
        const allSheets = spreadsheet.getSheets();
        for (const sheet of allSheets) {
          if (sheet.getName().match(/^PI\s*\d+\s*-\s*Capacity$/i)) {
            capacitySheet = sheet;
            break;
          }
        }
      }
    }
    
    if (!capacitySheet) {
      console.log('No consolidated capacity sheet found for role extraction');
      return null;
    }
    
    const dataRange = capacitySheet.getDataRange();
    const values = dataRange.getValues();
    const maxRows = values.length;
    const maxCols = values[0] ? values[0].length : 0;
    
    const normalizeTeamName = (name) => name.toUpperCase().replace(/[\s\-_]/g, '');
    const normalizedSearchTeam = normalizeTeamName(teamName);
    
    const teamBlockWidth = 11;
    const teamBlockHeight = 25;
    
    // Find EMA Clinical column
    let clinicalCol = -1;
    for (let col = 0; col < maxCols; col += teamBlockWidth) {
      const vsName = values[0][col];
      if (vsName && vsName.toString().toUpperCase().includes('EMA CLINICAL')) {
        clinicalCol = col;
        break;
      }
    }
    
    if (clinicalCol === -1) {
      console.log('EMA Clinical column not found');
      return null;
    }
    
    // Find the team
    let teamRow = -1;
    for (let row = 2; row < maxRows; row += teamBlockHeight) {
      const cellValue = values[row][clinicalCol];
      if (cellValue && normalizeTeamName(cellValue.toString()) === normalizedSearchTeam) {
        teamRow = row;
        break;
      }
    }
    
    if (teamRow === -1) {
      console.log(`Team "${teamName}" not found in EMA Clinical`);
      return null;
    }
    
    // Extract roles
    const roles = {};
    const totalCol = clinicalCol + 9;
    
    // Before FF roles (rows +11 to +16)
    for (let offset = 11; offset <= 16; offset++) {
      const row = teamRow + offset;
      if (row >= maxRows) break;
      
      const roleName = values[row][clinicalCol];
      const totalValue = values[row][totalCol];
      
      if (roleName && roleName.toString().trim() && 
          !roleName.toString().toLowerCase().includes('base capacity')) {
        const roleKey = roleName.toString().trim().toUpperCase();
        const normalizedRole = ROLE_NORMALIZATION_DANS[roleKey] || roleKey;
        
        let total = 0;
        if (totalValue && totalValue !== '-') {
          total = parseFloat(totalValue) || 0;
        }
        
        if (!roles[normalizedRole]) {
          roles[normalizedRole] = { displayName: roleName.toString().trim(), beforeFF: 0, afterFF: 0, total: 0 };
        }
        roles[normalizedRole].beforeFF = total;
      }
    }
    
    // After FF roles (rows +19 to +24)
    for (let offset = 19; offset <= 24; offset++) {
      const row = teamRow + offset;
      if (row >= maxRows) break;
      
      const roleName = values[row][clinicalCol];
      const totalValue = values[row][totalCol];
      
      if (roleName && roleName.toString().trim() && 
          !roleName.toString().toLowerCase().includes('base capacity')) {
        const roleKey = roleName.toString().trim().toUpperCase();
        const normalizedRole = ROLE_NORMALIZATION_DANS[roleKey] || roleKey;
        
        let total = 0;
        if (totalValue && totalValue !== '-') {
          total = parseFloat(totalValue) || 0;
        }
        
        if (!roles[normalizedRole]) {
          roles[normalizedRole] = { displayName: roleName.toString().trim(), beforeFF: 0, afterFF: 0, total: 0 };
        }
        roles[normalizedRole].afterFF = total;
      }
    }
    
    // Calculate combined totals
    Object.keys(roles).forEach(roleKey => {
      roles[roleKey].total = roles[roleKey].beforeFF + roles[roleKey].afterFF;
    });
    
    // Filter to only active roles
    const activeRoles = {};
    Object.keys(roles).forEach(roleKey => {
      if (roles[roleKey].total > 0) {
        activeRoles[roleKey] = roles[roleKey];
      }
    });
    
    return { roles: activeRoles };
    
  } catch (error) {
    console.error(`Error getting roles for team ${teamName}:`, error);
    return null;
  }
}

/**
 * Get role breakdown data from PI sheet for all clinical teams
 * @param {Sheet} piSheet - The PI data sheet
 * @param {Object} teamRoles - Object with roles per team from capacity sheet
 * @returns {Object} { teamName: { roleName: { total, byIteration } } }
 */
function getRoleBreakdownFromPISheet(piSheet, teamRoles) {
  if (!piSheet) return {};
  
  const dataRange = piSheet.getDataRange();
  const values = dataRange.getValues();
  
  if (values.length < 4) return {};
  
  const headers = values[3];
  const scrumTeamCol = headers.indexOf('Scrum Team');
  const storyPointsCol = headers.indexOf('Story Points');
  const issueTypeCol = headers.indexOf('Issue Type');
  const labelsCol = headers.indexOf('Labels');
  const summaryCol = headers.indexOf('Summary');
  const sprintCol = headers.indexOf('Sprint');
  
  if (scrumTeamCol === -1 || storyPointsCol === -1) return {};
  
  const roleData = {};
  
  // Initialize structure for all clinical teams
  DANS_REPORT_CONFIG.clinicalTeams.forEach(team => {
    roleData[team] = { Unassigned: { total: 0, byIteration: {1:0,2:0,3:0,4:0,5:0,6:0} } };
    
    const teamRoleInfo = teamRoles[team];
    if (teamRoleInfo && teamRoleInfo.roles) {
      Object.keys(teamRoleInfo.roles).forEach(role => {
        roleData[team][role] = { 
          total: 0, 
          byIteration: {1:0,2:0,3:0,4:0,5:0,6:0},
          capacity: teamRoleInfo.roles[role].total || 0,
          displayName: teamRoleInfo.roles[role].displayName || role
        };
      });
    }
  });
  
  // Process each row
  for (let i = 4; i < values.length; i++) {
    const row = values[i];
    const scrumTeam = row[scrumTeamCol];
    const storyPoints = parseFloat(row[storyPointsCol]) || 0;
    const issueType = issueTypeCol !== -1 ? row[issueTypeCol] : '';
    const labels = labelsCol !== -1 ? (row[labelsCol] || '').toString().split(',').map(l => l.trim()) : [];
    const summary = summaryCol !== -1 ? row[summaryCol] : '';
    const sprint = sprintCol !== -1 ? row[sprintCol] : '';
    
    if (!scrumTeam || storyPoints === 0) continue;
    if (issueType !== 'Story' && issueType !== 'Bug') continue;
    
    const cleanTeamName = scrumTeam.toString().trim().toUpperCase();
    
    if (!roleData[cleanTeamName]) continue;
    
    // Detect role
    let detectedRole = detectRoleFromTicketDans(labels, summary);
    
    // If role not in team's available roles, mark as unassigned
    if (detectedRole && !roleData[cleanTeamName][detectedRole]) {
      detectedRole = null;
    }
    
    const roleKey = detectedRole || 'Unassigned';
    
    if (!roleData[cleanTeamName][roleKey]) {
      roleData[cleanTeamName][roleKey] = { total: 0, byIteration: {1:0,2:0,3:0,4:0,5:0,6:0}, capacity: 0 };
    }
    
    roleData[cleanTeamName][roleKey].total += Math.ceil(storyPoints);
    
    // Extract iteration from sprint name
    const sprintMatch = sprint.match(/\d+\s*\.\s*(\d)/);
    if (sprintMatch) {
      const iteration = parseInt(sprintMatch[1]);
      if (iteration >= 1 && iteration <= 6) {
        roleData[cleanTeamName][roleKey].byIteration[iteration] += Math.ceil(storyPoints);
      }
    }
  }
  
  return roleData;
}

/**
 * Create Role Breakdown section for Dan's Report
 * @param {Sheet} sheet - The report sheet
 * @param {number} startRow - Starting row
 * @param {Spreadsheet} spreadsheet - The spreadsheet object
 * @param {Sheet} piSheet - The PI data sheet
 * @returns {number} Next available row
 */
function createDansReportRoleBreakdown(sheet, startRow, spreadsheet, piSheet) {
  console.log('Creating Role Breakdown section for Dan\'s Report');
  
  // Get roles for each clinical team from capacity sheet
  const teamRoles = {};
  let hasAnyRoles = false;
  
  DANS_REPORT_CONFIG.clinicalTeams.forEach(team => {
    const roles = getClinicalTeamRolesFromCapacity(spreadsheet, team);
    if (roles && roles.roles && Object.keys(roles.roles).length > 0) {
      teamRoles[team] = roles;
      hasAnyRoles = true;
    }
  });
  
  if (!hasAnyRoles) {
    console.log('No role data found in capacity sheet - skipping Role Breakdown section');
    return startRow;
  }
  
  // Get role breakdown from PI sheet
  const roleBreakdown = getRoleBreakdownFromPISheet(piSheet, teamRoles);
  
  // Section title
  sheet.getRange(startRow, 1).setValue('ROLE BREAKDOWN BY TEAM');
  sheet.getRange(startRow, 1, 1, 12)
    .setBackground('#4A235A')
    .setFontColor('white')
    .setFontSize(14)
    .setFontWeight('bold')
    .setFontFamily('Comfortaa');
  sheet.getRange(startRow, 1, 1, 12).merge();
  startRow += 2;
  
  // Create a table for each team that has role data
  DANS_REPORT_CONFIG.clinicalTeams.forEach(teamName => {
    if (!teamRoles[teamName]) return;
    
    const teamRoleBreakdown = roleBreakdown[teamName] || {};
    const availableRoles = teamRoles[teamName].roles || {};
    
    // Check if there's any data to show
    const rolesToShow = Object.keys(availableRoles).filter(r => 
      (availableRoles[r].total > 0) || (teamRoleBreakdown[r] && teamRoleBreakdown[r].total > 0)
    );
    
    // Add Unassigned if it has data
    if (teamRoleBreakdown['Unassigned'] && teamRoleBreakdown['Unassigned'].total > 0) {
      rolesToShow.push('Unassigned');
    }
    
    if (rolesToShow.length === 0) return;
    
    // Sort: named roles by capacity desc, Unassigned last
    rolesToShow.sort((a, b) => {
      if (a === 'Unassigned') return 1;
      if (b === 'Unassigned') return -1;
      return (availableRoles[b]?.total || 0) - (availableRoles[a]?.total || 0);
    });
    
    // Team header
    sheet.getRange(startRow, 1).setValue(teamName);
    sheet.getRange(startRow, 1, 1, 12)
      .setBackground('#9b7bb8')
      .setFontColor('white')
      .setFontSize(10)
      .setFontWeight('bold')
      .setFontFamily('Comfortaa');
    startRow++;
    
    // Column headers
    const headers = ['Role', 'Baseline', 'Used', 'Remaining', '% Rem'];
    sheet.getRange(startRow, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(startRow, 1, 1, headers.length)
      .setBackground('#d5c4e1')
      .setFontWeight('bold')
      .setFontSize(8)
      .setFontFamily('Comfortaa')
      .setHorizontalAlignment('center');
    startRow++;
    
    // Data rows
    rolesToShow.forEach(roleKey => {
      const roleCapacity = availableRoles[roleKey]?.total || 0;
      const roleUsed = teamRoleBreakdown[roleKey]?.total || 0;
      const remaining = roleCapacity - roleUsed;
      const percentRemaining = roleCapacity > 0 ? Math.round((remaining / roleCapacity) * 100) : (roleUsed > 0 ? -100 : 0);
      
      const displayName = availableRoles[roleKey]?.displayName || roleKey;
      
      const rowData = [
        displayName,
        Math.ceil(roleCapacity),
        Math.ceil(roleUsed),
        Math.ceil(remaining),
        percentRemaining + '%'
      ];
      
      sheet.getRange(startRow, 1, 1, rowData.length).setValues([rowData]);
      sheet.getRange(startRow, 1, 1, headers.length)
        .setFontSize(8)
        .setFontFamily('Comfortaa');
      sheet.getRange(startRow, 2, 1, 4).setHorizontalAlignment('center');
      
      // Color code Remaining and % Remaining
      if (remaining >= 0) {
        sheet.getRange(startRow, 4).setBackground('#ccffcc');
        sheet.getRange(startRow, 5).setBackground('#ccffcc');
      } else {
        sheet.getRange(startRow, 4).setBackground('#ffcccc');
        sheet.getRange(startRow, 5).setBackground('#ffcccc');
      }
      
      // Highlight Unassigned
      if (roleKey === 'Unassigned' && roleUsed > 0) {
        sheet.getRange(startRow, 1).setBackground('#fff3cd');
        sheet.getRange(startRow, 3).setBackground('#fff3cd');
      }
      
      startRow++;
    });
    
    startRow++; // Space between teams
  });
  
  return startRow;
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
 * - Planned Load (FP): Sum of Feature Points x 10
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
  console.log('Planned Load: Sum(Feature Points) x 10');
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
  
  const anticipatedLoad = {}; // Feature Points x 10
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
    
    // Add Feature Points x 10 if > 0
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
  
  console.log('\n=== Planned Load Summary (FP x 10) ===');
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
    'Capacity Remaining for Use',
    'Planned Load (FP)',
    'Planned Remaining for Use',
    'Actual Load (SPE)',
    'Actual Remaining for Use',
    'Baseline Capacity',
    'Capacity Used (LOE)',
    'Capacity Remaining for Use',
    'Planned Load (FP)',
    'Planned Remaining for Use',
    'Actual Load (SPE)',
    'Actual Remaining for Use'
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
  sheet.getRange(currentRow, 6).setNote('At Epic level: sum of (Feature Points x 10)\nFiltered by Product - Feature or Product - Compliance allocation');
  
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
  sheet.getRange(currentRow, 6).setNote('Feature Points at Epic level\nUsed to calculate Planned Load (FP x 10)');
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

/**
 * Create the capacity utilization report with both tables
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