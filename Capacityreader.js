/**
 * CapacityReader.js
 *
 * Reads capacity data from the new consolidated capacity planning format.
 *
 * NEW FORMAT STRUCTURE:
 * - Single tab contains all value streams and teams
 * - Row 1: Value Stream names at columns 1, 12, 23, 34, 45, etc. (every 11 columns)
 * - Row 3: First scrum team row (teams repeat every 25 rows: 3, 28, 53, 78, etc.)
 * - Each team block is 11 columns wide and 25 rows tall
 *
 * TEAM BLOCK STRUCTURE (relative to team start row N and column C):
 * - Row N (C): Team name
 * - Row N+1 (C): "Allocation Type" header
 * - Row N+2 (C+2): KLO, values in C+3 to C+8, Total in C+9
 * - Row N+3 (C+2): Quality
 * - Row N+4 (C+2): Tech / Platform
 * - Row N+5 (C+2): Product - Feature
 * - Row N+6 (C+2): Product - Compliance
 * - Row N+7 (C+2): Unplanned work
 * - Row N+8 (C+3 to C+9): Totals per iteration and grand total
 * - Row N+10 (C): "Base Capacity before FF" header
 * - Row N+16 (C+3 to C+9): Capacity totals before Feature Freeze
 * - Row N+18 (C): "Base Capacity after FF" header
 * - Row N+24 (C+3 to C+9): Capacity totals after Feature Freeze
 */

// Configuration for the new capacity format
// IMPORTANT: Update sheetName to match your actual capacity planning tab name
const CAPACITY_FORMAT_CONFIG = {
  // Primary sheet name pattern - will be dynamically replaced with current PI number
  // Example: "PI14 - Capacity", "PI15 - Capacity"
  sheetNamePattern: 'PI{PI_NUMBER} - Capacity',

  // Alternative sheet names to try if primary not found
  alternativeSheetNames: ['Capacity Planning', 'Consolidated Capacity'],

  teamBlockWidth: 11,              // Each team block is 11 columns wide
  teamBlockHeight: 25,             // Each team block is 25 rows tall
  valueStreamRow: 1,               // Row containing value stream names
  firstTeamRow: 3,                 // First row containing team names

  // Relative row offsets within a team block (from team name row)
  relativeRows: {
    teamName: 0,
    allocationHeader: 1,
    klo: 2,
    quality: 3,
    techPlatform: 4,
    productFeature: 5,
    productCompliance: 6,
    unplannedWork: 7,
    allocationTotal: 8,
    baseCapacityBeforeFFHeader: 10,
    baseCapacityBeforeFFTotal: 16,
    baseCapacityAfterFFHeader: 18,
    baseCapacityAfterFFTotal: 24
  },

  // Relative column offsets within a team block (from team start column, 0-indexed)
  relativeColumns: {
    teamName: 0,        // Column 1 of block
    allocationLabel: 2, // Column 3 of block (KLO, Quality, etc.)
    iteration1: 3,      // Column 4 (14.1)
    iteration2: 4,      // Column 5 (14.2)
    iteration3: 5,      // Column 6 (14.3)
    iteration4: 6,      // Column 7 (14.4)
    iteration5: 7,      // Column 8 (14.5)
    iteration6: 8,      // Column 9 (14.6)
    total: 9            // Column 10 (T)
  }
};

/**
 * Find the capacity planning sheet, trying primary and alternative names
 * @param {Spreadsheet} spreadsheet - The spreadsheet object
 * @param {number|string} piNumber - Optional PI number for dynamic sheet name
 * @returns {Sheet|null} The capacity sheet or null if not found
 */
function findCapacityPlanningSheet(spreadsheet, piNumber = null) {
  const config = CAPACITY_FORMAT_CONFIG;

  // If PI number provided, try the PI-specific sheet name first
  if (piNumber) {
    const piSheetName = config.sheetNamePattern.replace('{PI_NUMBER}', piNumber);
    let sheet = spreadsheet.getSheetByName(piSheetName);
    if (sheet) {
      console.log(`Found capacity sheet: "${piSheetName}"`);
      return sheet;
    }
    console.log(`PI-specific sheet "${piSheetName}" not found, trying alternatives...`);
  }

  // Try to auto-detect PI number from existing sheets
  const allSheets = spreadsheet.getSheets();
  for (const sheet of allSheets) {
    const name = sheet.getName();
    // Match pattern like "PI14 - Capacity", "PI15 - Capacity"
    const match = name.match(/^PI\s*(\d+)\s*-\s*Capacity$/i);
    if (match) {
      console.log(`Found capacity sheet by pattern: "${name}" (PI ${match[1]})`);
      return sheet;
    }
  }

  // Try alternative names
  for (const altName of config.alternativeSheetNames) {
    const sheet = spreadsheet.getSheetByName(altName);
    if (sheet) {
      console.log(`Found capacity sheet using alternative name: "${altName}"`);
      return sheet;
    }
  }

  console.log('No consolidated capacity sheet found');
  return null;
}

/**
 * Find all teams in the consolidated capacity sheet
 * @param {Sheet} capacitySheet - The capacity planning sheet
 * @returns {Array} Array of {teamName, valueStream, startRow, startCol}
 */
function findAllTeamsInCapacitySheet(capacitySheet) {
  const config = CAPACITY_FORMAT_CONFIG;
  const teams = [];

  const dataRange = capacitySheet.getDataRange();
  const values = dataRange.getValues();
  const maxRows = values.length;
  const maxCols = values[0] ? values[0].length : 0;

  console.log(`Scanning capacity sheet: ${maxRows} rows x ${maxCols} columns`);

  // Scan for value streams in row 1
  const valueStreams = [];
  for (let col = 0; col < maxCols; col += config.teamBlockWidth) {
    const vsName = values[config.valueStreamRow - 1][col];
    if (vsName && vsName.toString().trim()) {
      valueStreams.push({
        name: vsName.toString().trim(),
        startCol: col
      });
      console.log(`Found Value Stream: "${vsName}" at column ${col + 1}`);
    }
  }

  // For each value stream column, scan down for team names
  valueStreams.forEach(vs => {
    let row = config.firstTeamRow - 1; // Convert to 0-indexed

    while (row < maxRows) {
      const teamName = values[row][vs.startCol];

      if (teamName && teamName.toString().trim()) {
        const teamStr = teamName.toString().trim();

        // Skip header rows
        if (!teamStr.toLowerCase().includes('allocation') &&
            !teamStr.toLowerCase().includes('base capacity') &&
            !teamStr.toLowerCase().includes('mobile') &&
            !teamStr.toLowerCase().includes('qa') &&
            !teamStr.toLowerCase().includes('w-dev') &&
            !teamStr.toLowerCase().includes('m-dev') &&
            !teamStr.toLowerCase().includes('be') &&
            !teamStr.toLowerCase().includes('fe') &&
            !teamStr.toLowerCase().includes('aqa')) {

          teams.push({
            teamName: teamStr,
            valueStream: vs.name,
            startRow: row + 1, // Convert back to 1-indexed for sheet operations
            startCol: vs.startCol + 1 // Convert back to 1-indexed
          });
          console.log(`  Found Team: "${teamStr}" at row ${row + 1}, col ${vs.startCol + 1}`);
        }
      }

      row += config.teamBlockHeight;
    }
  });

  console.log(`Total teams found: ${teams.length}`);
  return teams;
}

/**
 * Find a specific team in the capacity sheet
 * @param {Sheet} capacitySheet - The capacity planning sheet
 * @param {string} teamName - Team name to find
 * @param {string} valueStream - Optional value stream to filter by
 * @returns {Object|null} {teamName, valueStream, startRow, startCol} or null if not found
 */
function findTeamInCapacitySheet(capacitySheet, teamName, valueStream = null) {
  const teams = findAllTeamsInCapacitySheet(capacitySheet);

  // Normalize team name for comparison
  const normalizeTeamName = (name) => {
    return name.toUpperCase().replace(/[\s\-_]/g, '');
  };

  const normalizedSearch = normalizeTeamName(teamName);

  // First try exact match (case-insensitive)
  let match = teams.find(t => {
    const matches = normalizeTeamName(t.teamName) === normalizedSearch;
    if (valueStream && matches) {
      return t.valueStream.toUpperCase().includes(valueStream.toUpperCase());
    }
    return matches;
  });

  // If not found, try partial match
  if (!match) {
    match = teams.find(t => {
      const normalizedTeam = normalizeTeamName(t.teamName);
      const partialMatch = normalizedTeam.includes(normalizedSearch) ||
                          normalizedSearch.includes(normalizedTeam);
      if (valueStream && partialMatch) {
        return t.valueStream.toUpperCase().includes(valueStream.toUpperCase());
      }
      return partialMatch;
    });
  }

  if (match) {
    console.log(`Found team "${teamName}" -> "${match.teamName}" in ${match.valueStream} at row ${match.startRow}, col ${match.startCol}`);
  } else {
    console.log(`Team "${teamName}" not found in capacity sheet`);
  }

  return match;
}

/**
 * Get capacity data for a specific team from the consolidated format
 * @param {Spreadsheet} spreadsheet - The spreadsheet object
 * @param {string} teamName - Team name to get capacity for
 * @param {string} valueStream - Optional value stream filter
 * @returns {Object|null} Capacity data object
 */
function getCapacityDataForTeamConsolidated(spreadsheet, teamName, valueStream = null) {
  try {
    const config = CAPACITY_FORMAT_CONFIG;
    const capacitySheet = findCapacityPlanningSheet(spreadsheet);

    if (!capacitySheet) {
      console.log('No consolidated capacity sheet found');
      return null;
    }

    const teamLocation = findTeamInCapacitySheet(capacitySheet, teamName, valueStream);
    if (!teamLocation) {
      return null;
    }

    const startRow = teamLocation.startRow;
    const startCol = teamLocation.startCol;
    const relRows = config.relativeRows;
    const relCols = config.relativeColumns;

    // Helper function to get cell value
    const getCellValue = (rowOffset, colOffset) => {
      const cell = capacitySheet.getRange(startRow + rowOffset, startCol + colOffset);
      const value = cell.getValue();
      // Handle '-' or empty as 0
      if (value === '-' || value === '' || value === null) return 0;
      return parseFloat(value) || 0;
    };

    // Read allocation data (totals column)
    const allocations = {
      klo: getCellValue(relRows.klo, relCols.total),
      quality: getCellValue(relRows.quality, relCols.total),
      techPlatform: getCellValue(relRows.techPlatform, relCols.total),
      productFeature: getCellValue(relRows.productFeature, relCols.total),
      productCompliance: getCellValue(relRows.productCompliance, relCols.total),
      unplannedWork: getCellValue(relRows.unplannedWork, relCols.total)
    };

    // Read iteration-level data
    const iterationData = {};
    for (let iter = 1; iter <= 6; iter++) {
      const colOffset = relCols.iteration1 + (iter - 1);
      iterationData[iter] = {
        klo: getCellValue(relRows.klo, colOffset),
        quality: getCellValue(relRows.quality, colOffset),
        techPlatform: getCellValue(relRows.techPlatform, colOffset),
        productFeature: getCellValue(relRows.productFeature, colOffset),
        productCompliance: getCellValue(relRows.productCompliance, colOffset),
        unplannedWork: getCellValue(relRows.unplannedWork, colOffset)
      };
    }

    // Read base capacity totals
    const baseCapacityBeforeFF = getCellValue(relRows.baseCapacityBeforeFFTotal, relCols.total);
    const baseCapacityAfterFF = getCellValue(relRows.baseCapacityAfterFFTotal, relCols.total);

    // Calculate totals
    const totalAllocation = allocations.klo + allocations.quality + allocations.techPlatform +
                           allocations.productFeature + allocations.productCompliance + allocations.unplannedWork;

    // Product capacity = Product Feature + Product Compliance
    const productCapacity = allocations.productFeature + allocations.productCompliance;

    const capacityData = {
      teamName: teamLocation.teamName,
      valueStream: teamLocation.valueStream,

      // Allocation breakdown
      allocations: allocations,

      // Iteration-level data
      byIteration: iterationData,

      // Summary totals
      total: totalAllocation,
      productCapacity: productCapacity,
      baseCapacityBeforeFF: baseCapacityBeforeFF,
      baseCapacityAfterFF: baseCapacityAfterFF,

      // For compatibility with existing code
      'Features': productCapacity,
      'Tech/Platform': allocations.techPlatform,
      'Planned KLO': allocations.klo,
      'Planned Quality': allocations.quality,
      'Unplanned': allocations.unplannedWork
    };

    console.log(`Capacity data for ${teamName}:`, JSON.stringify(capacityData, null, 2));
    return capacityData;

  } catch (error) {
    console.error(`Error getting capacity data for team ${teamName}:`, error);
    return null;
  }
}

/**
 * Get capacity data for all teams in a value stream from the consolidated format
 * @param {Spreadsheet} spreadsheet - The spreadsheet object
 * @param {Array} issues - Array of issues to filter teams by
 * @param {string} valueStream - Value stream to get capacity for
 * @returns {Object|null} Capacity data object compatible with getCapacityDataDynamic
 */
function getCapacityDataDynamicConsolidated(spreadsheet, issues, valueStream) {
  try {
    const config = CAPACITY_FORMAT_CONFIG;
    const capacitySheet = findCapacityPlanningSheet(spreadsheet);

    if (!capacitySheet) {
      console.log('No consolidated capacity sheet found, falling back to legacy format');
      // Fall back to legacy function if it exists
      if (typeof getCapacityDataDynamic === 'function') {
        return getCapacityDataDynamic(spreadsheet, issues, valueStream);
      }
      return null;
    }

    console.log(`\n=== Getting Consolidated Capacity Data for Value Stream: ${valueStream} ===`);

    // Get all teams from the capacity sheet
    const allCapacityTeams = findAllTeamsInCapacitySheet(capacitySheet);

    // Filter teams by value stream if specified
    let relevantTeams = allCapacityTeams;
    if (valueStream && valueStream.trim()) {
      relevantTeams = allCapacityTeams.filter(t =>
        t.valueStream.toUpperCase().includes(valueStream.toUpperCase())
      );
      console.log(`Filtered to ${relevantTeams.length} teams for value stream "${valueStream}"`);
    }

    // Get unique scrum teams from issues
    const issueTeams = [...new Set(issues.map(issue => issue.scrumTeam || 'Unassigned'))];
    console.log(`Teams in issues: ${issueTeams.join(', ')}`);

    // Normalize team names for matching
    const normalizeTeamName = (name) => name.toUpperCase().replace(/[\s\-_]/g, '');

    // Build capacity data structure
    const capacityData = {
      byTeam: {},
      byAllocation: {
        'Features': 0,
        'Tech/Platform': 0,
        'Planned KLO': 0,
        'Planned Quality': 0,
        'Unplanned': 0
      },
      total: 0
    };

    // Process each relevant team
    relevantTeams.forEach(teamInfo => {
      const teamCapacity = getCapacityDataForTeamConsolidated(spreadsheet, teamInfo.teamName, valueStream);

      if (teamCapacity) {
        // Check if this team is in our issues
        const normalizedCapacityTeam = normalizeTeamName(teamInfo.teamName);
        const matchingIssueTeam = issueTeams.find(t =>
          normalizeTeamName(t) === normalizedCapacityTeam ||
          normalizeTeamName(t).includes(normalizedCapacityTeam) ||
          normalizedCapacityTeam.includes(normalizeTeamName(t))
        );

        if (matchingIssueTeam || !issues.length) {
          const teamKey = matchingIssueTeam || teamInfo.teamName;

          capacityData.byTeam[teamKey] = {
            'Features': teamCapacity.productCapacity,
            'Tech/Platform': teamCapacity.allocations.techPlatform,
            'Planned KLO': teamCapacity.allocations.klo,
            'Planned Quality': teamCapacity.allocations.quality,
            'Unplanned': teamCapacity.allocations.unplannedWork,
            total: teamCapacity.total,
            baseCapacityBeforeFF: teamCapacity.baseCapacityBeforeFF,
            baseCapacityAfterFF: teamCapacity.baseCapacityAfterFF,
            byIteration: teamCapacity.byIteration
          };

          // Accumulate totals
          capacityData.byAllocation['Features'] += teamCapacity.productCapacity;
          capacityData.byAllocation['Tech/Platform'] += teamCapacity.allocations.techPlatform;
          capacityData.byAllocation['Planned KLO'] += teamCapacity.allocations.klo;
          capacityData.byAllocation['Planned Quality'] += teamCapacity.allocations.quality;
          capacityData.byAllocation['Unplanned'] += teamCapacity.allocations.unplannedWork;
          capacityData.total += teamCapacity.total;

          console.log(`  ✓ Team "${teamKey}": ${teamCapacity.total} total capacity`);
        } else {
          console.log(`  ✗ Team "${teamInfo.teamName}": Not in issues, skipping`);
        }
      }
    });

    console.log(`\n=== Capacity Summary ===`);
    console.log(`Total capacity: ${capacityData.total} points`);
    console.log(`Teams included (${Object.keys(capacityData.byTeam).length}): ${Object.keys(capacityData.byTeam).join(', ')}`);
    console.log(`Allocation breakdown:`, capacityData.byAllocation);
    console.log('=== End Consolidated Capacity Data ===\n');

    return capacityData;

  } catch (error) {
    console.error('Error reading consolidated capacity data:', error);
    console.error('Stack trace:', error.stack);
    return null;
  }
}

/**
 * Get iteration-level capacity data for a team
 * @param {Spreadsheet} spreadsheet - The spreadsheet object
 * @param {string} teamName - Team name
 * @param {number} iteration - Iteration number (1-6)
 * @param {string} valueStream - Optional value stream filter
 * @returns {Object|null} Iteration capacity data
 */
function getIterationCapacityForTeam(spreadsheet, teamName, iteration, valueStream = null) {
  const teamCapacity = getCapacityDataForTeamConsolidated(spreadsheet, teamName, valueStream);

  if (!teamCapacity || !teamCapacity.byIteration || !teamCapacity.byIteration[iteration]) {
    return null;
  }

  return teamCapacity.byIteration[iteration];
}

/**
 * Get base capacity (before/after Feature Freeze) for a team
 * @param {Spreadsheet} spreadsheet - The spreadsheet object
 * @param {string} teamName - Team name
 * @param {boolean} beforeFF - True for before Feature Freeze, false for after
 * @param {string} valueStream - Optional value stream filter
 * @returns {number} Base capacity value
 */
function getBaseCapacityForTeam(spreadsheet, teamName, beforeFF = true, valueStream = null) {
  const teamCapacity = getCapacityDataForTeamConsolidated(spreadsheet, teamName, valueStream);

  if (!teamCapacity) {
    return 0;
  }

  return beforeFF ? teamCapacity.baseCapacityBeforeFF : teamCapacity.baseCapacityAfterFF;
}

/**
 * Check if the new consolidated capacity format is available
 * @param {Spreadsheet} spreadsheet - The spreadsheet object
 * @returns {boolean} True if consolidated format sheet exists
 */
function hasConsolidatedCapacityFormat(spreadsheet) {
  const sheet = findCapacityPlanningSheet(spreadsheet);
  return sheet !== null;
}

/**
 * Get capacity data using the appropriate format (consolidated or legacy)
 * This is the main entry point that auto-detects the format
 * @param {Spreadsheet} spreadsheet - The spreadsheet object
 * @param {Array} issues - Array of issues
 * @param {string} valueStream - Value stream name
 * @returns {Object|null} Capacity data
 */
function getCapacityDataAuto(spreadsheet, issues, valueStream) {
  // Check for consolidated format first
  if (hasConsolidatedCapacityFormat(spreadsheet)) {
    console.log('Using consolidated capacity format');
    return getCapacityDataDynamicConsolidated(spreadsheet, issues, valueStream);
  }

  // Fall back to legacy format
  console.log('Using legacy capacity format');
  if (typeof getCapacityDataDynamic === 'function') {
    return getCapacityDataDynamic(spreadsheet, issues, valueStream);
  }

  console.log('No capacity data function available');
  return null;
}

/**
 * Get capacity data for a single team using the appropriate format
 * @param {Spreadsheet} spreadsheet - The spreadsheet object
 * @param {string} teamName - Team name
 * @returns {Object|null} Capacity data for the team
 */
function getCapacityForTeamAuto(spreadsheet, teamName) {
  // Check for consolidated format first
  if (hasConsolidatedCapacityFormat(spreadsheet)) {
    console.log(`Using consolidated capacity format for team: ${teamName}`);
    return getCapacityDataForTeamConsolidated(spreadsheet, teamName);
  }

  // Fall back to legacy format
  console.log(`Using legacy capacity format for team: ${teamName}`);
  if (typeof getCapacityDataForTeam === 'function') {
    return getCapacityDataForTeam(spreadsheet, teamName);
  }

  console.log('No team capacity function available');
  return null;
}

// ===== DIAGNOSTIC FUNCTIONS =====

/**
 * Test function to validate the capacity sheet structure
 */
function testCapacitySheetStructure() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const config = CAPACITY_FORMAT_CONFIG;

  console.log('\n========================================');
  console.log('CAPACITY SHEET STRUCTURE TEST');
  console.log('========================================\n');

  // Check for consolidated format
  const consolidatedSheet = findCapacityPlanningSheet(spreadsheet);
  if (consolidatedSheet) {
    console.log(`✓ Found consolidated capacity sheet: "${consolidatedSheet.getName()}"`);

    const teams = findAllTeamsInCapacitySheet(consolidatedSheet);
    console.log(`\nTeams found: ${teams.length}`);

    // Group by value stream
    const byVS = {};
    teams.forEach(t => {
      if (!byVS[t.valueStream]) byVS[t.valueStream] = [];
      byVS[t.valueStream].push(t.teamName);
    });

    Object.keys(byVS).forEach(vs => {
      console.log(`\n${vs}:`);
      byVS[vs].forEach(team => console.log(`  - ${team}`));
    });

    // Test reading data for first team
    if (teams.length > 0) {
      console.log(`\n\nTesting data read for first team: "${teams[0].teamName}"`);
      const teamData = getCapacityDataForTeamConsolidated(spreadsheet, teams[0].teamName);
      if (teamData) {
        console.log('✓ Successfully read team data');
        console.log(`  Total capacity: ${teamData.total}`);
        console.log(`  Product capacity: ${teamData.productCapacity}`);
        console.log(`  Base capacity before FF: ${teamData.baseCapacityBeforeFF}`);
      } else {
        console.log('✗ Failed to read team data');
      }
    }
  } else {
    console.log(`✗ No consolidated capacity sheet found`);
    console.log(`  Tried: "${config.sheetName}" and alternatives: ${config.alternativeSheetNames.join(', ')}`);
  }

  // Check for legacy format
  const legacySheet = spreadsheet.getSheetByName('Capacity');
  if (legacySheet) {
    console.log(`\n✓ Found legacy capacity sheet: "Capacity"`);
  }

  console.log('\n========================================');
  console.log('END CAPACITY SHEET STRUCTURE TEST');
  console.log('========================================\n');
}

/**
 * Menu function to run the capacity test
 */
function menuTestCapacityStructure() {
  testCapacitySheetStructure();
  SpreadsheetApp.getUi().alert('Capacity structure test complete. Check the logs (View > Logs) for details.');
}
