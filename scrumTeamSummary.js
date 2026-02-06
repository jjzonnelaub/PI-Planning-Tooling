// ===== UTILITY FUNCTIONS =====
function setRowHeightWithLimit(sheet, row, desiredHeight, maxHeight = 70) {
  sheet.setRowHeight(row, Math.min(desiredHeight, maxHeight));
}

function parseCostOfDelay(value) {
  if (!value) return 0;
  
  // If already a number, return it
  if (typeof value === 'number') return value;
  
  // Convert to string and clean up
  let cleanValue = value.toString().trim();
  
  // Remove currency symbols and commas
  cleanValue = cleanValue.replace(/[$,]/g, '');
  
  // Handle shorthand notations (1M, 500K, etc.)
  if (cleanValue.match(/(\d+\.?\d*)([KMB])/i)) {
    const match = cleanValue.match(/(\d+\.?\d*)([KMB])/i);
    const num = parseFloat(match[1]);
    const multiplier = match[2].toUpperCase();
    
    switch (multiplier) {
      case 'K': return num * 1000;
      case 'M': return num * 1000000;
      case 'B': return num * 1000000000;
    }
  }
  
  // Try to parse as regular number
  const parsed = parseFloat(cleanValue);
  return isNaN(parsed) ? 0 : parsed;
}

function createProgressBar(sheet, row, column, percentage) {
  const barLength = 20;
  const filledLength = Math.round((percentage / 100) * barLength);
  const emptyLength = barLength - filledLength;
  
  const progressBar = '#'.repeat(filledLength) + '-'.repeat(emptyLength);
  
  sheet.getRange(row, column).setValue(progressBar);
  sheet.getRange(row, column).setFontFamily('Courier New');
  
  // Color based on percentage
  let color;
  if (percentage >= 80) {
    color = '#4CAF50'; // Green
  } else if (percentage >= 60) {
    color = '#FFC107'; // Yellow
  } else {
    color = '#F44336'; // Red
  }
  
  sheet.getRange(row, column).setFontColor(color);
}

function showProgress(message) {
  const template = HtmlService.createHtmlOutput(`
    <div style="padding: 20px; text-align: center;">
      <p>${message}</p>
      <div style="margin-top: 10px;">
        <div style="display: inline-block; border: 1px solid #ccc; width: 200px; height: 20px;">
          <div style="background: #4CAF50; width: 50%; height: 100%; animation: progress 2s infinite;"></div>
        </div>
      </div>
    </div>
    <style>
      @keyframes progress {
        0% { width: 0%; }
        50% { width: 100%; }
        100% { width: 0%; }
      }
    </style>
  `);
  
  SpreadsheetApp.getUi().showModalDialog(template, 'Processing...');
}

function closeProgress() {
  const html = HtmlService.createHtmlOutput('<script>google.script.host.close();</script>');
  SpreadsheetApp.getUi().showModalDialog(html, 'Closing...');
}

// ===== DATA PARSING FUNCTIONS =====
function parsePISheetRow(row, headers) {
  if (!row || !headers) {
    console.error('Missing row or headers in parsePISheetRow');
    return null;
  }
  
  const columnIndices = {};
  headers.forEach((header, index) => {
    if (header) {
      columnIndices[header] = index;
    }
  });
  
  // Helper function to safely get and clean value
  const getValue = (columnName, defaultValue = '') => {
    const index = columnIndices[columnName];
    const rawValue = (index !== undefined && row[index] !== undefined) ? row[index] : defaultValue;
    
    // Clean stringified objects
    if (typeof rawValue === 'string' && rawValue.includes('{') && rawValue.includes('value=')) {
      return parseSheetCellValue(rawValue);
    }
    
    return rawValue;
  };
  
  // Helper function to safely get numeric value
  const getNumericValue = (columnName, defaultValue = 0) => {
    const value = getValue(columnName);
    const num = Number(value);
    return isNaN(num) ? defaultValue : num;
  };
  
  // Parse ALL fields with cleaning
  return {
    key: getValue('Key'),
    parentKey: getValue('Parent Key'),
    epicLink: getValue('Epic Link'),
    issueType: getValue('Issue Type'),
    summary: getValue('Summary'),
    status: getValue('Status'),
    valueStream: getValue('Value Stream'),
    // ⭐ FIX: Properly set analyzedValueStream with fallback (this was being overwritten!)
    analyzedValueStream: getValue('Analyzed Value Stream') || getValue('Value Stream') || 'Unknown',
    org: getValue('Org'),
    piCommitment: getValue('PI Commitment'),
    programIncrement: getValue('Program Increment'),
    scrumTeam: getValue('Scrum Team'),
    piTargetIteration: getValue('PI Target Iteration'),
    iterationStart: getValue('Iteration Start'),
    iterationEnd: getValue('Iteration End'),
    allocation: getValue('Allocation'),
    portfolioInitiative: getValue('Portfolio Initiative'),
    programInitiative: getValue('Program Initiative'),
    rag: getValue('RAG'),
    ragNote: getValue('RAG Note'),
    storyPoints: getNumericValue('Story Points'),
    storyPointEstimate: getNumericValue('Story Point Estimate'),
    featurePoints: getNumericValue('Feature Points'),
    loeEstimate: getNumericValue('LOE Estimate'),
    // ⭐ REMOVED: Don't overwrite analyzedValueStream here (was line 549)
    properAllocation: getValue('Proper Allocation'),
    rowLastUpdated: getValue('Row Last Updated'),
    dependsOnValuestream: getValue('Depends on Valuestream'),
    costOfDelay: getNumericValue('Cost of Delay'),
    components: getValue('Components'),
    closedTransitionDate: getValue('Closed Transition Date'),
    workType: getValue('Work Type'),
    momentum: getValue('Momentum'),
    sprintName: getValue('Sprint Name'),
    fixVersion: getValue('Fix Version')
  };
}

function createValueStreamPlanningProgress(sheet, startRow, allIssues, epics, stories, valueStream) {
  console.log(`Creating planning progress section for ${valueStream}`);
  
  const spreadsheet = sheet.getParent();
  const capacityData = getCapacityDataDynamic(spreadsheet, allIssues, valueStream);
  
  const epicStories = stories.filter(s => s.issueType === 'Story' && (s.epicLink || s.parentKey));
  const epicsWithAllStoryPoints = new Set();
  
  epics.forEach(epic => {
    const epicChildStories = epicStories.filter(s => 
      (s.parentKey === epic.key || s.epicLink === epic.key) &&
      s.issueType === 'Story'
    );
    
    if (epicChildStories.length > 0 && epicChildStories.every(s => s.storyPoints && s.storyPoints > 0)) {
      epicsWithAllStoryPoints.add(epic.key);
    }
  });
  
  const percentEpicsWithStoryPoints = epics.length > 0 ? 
    Math.round((epicsWithAllStoryPoints.size / epics.length) * 100) : 0;
  
  let percentCapacityAllocated = 0;
  let capacityStatus = 'N/A';
  if (capacityData && capacityData.total > 0) {
    const totalStoryPoints = stories.reduce((sum, s) => sum + (s.storyPoints || 0), 0);
    percentCapacityAllocated = Math.round((totalStoryPoints / capacityData.total) * 100);
    capacityStatus = `${totalStoryPoints} / ${capacityData.total} points`;
  }
  
  sheet.getRange(startRow, 1).setValue('Planning Progress');
  sheet.getRange(startRow, 1, 1, 4).setBackground('#E1D5E7');
  sheet.getRange(startRow, 1).setFontSize(14).setFontWeight('bold').setFontColor('black');
  sheet.getRange(startRow, 1).setFontFamily('Comfortaa');
  sheet.getRange(startRow, 1).setVerticalAlignment('middle');
  startRow++;
  
  const metricsHeaders = ['Metric', '', 'Value', 'Progress'];
  sheet.getRange(startRow, 1, 1, metricsHeaders.length).setValues([metricsHeaders]);
  sheet.getRange(startRow, 1, 1, metricsHeaders.length)
    .setFontWeight('bold')
    .setBackground('#9b7bb8')
    .setFontColor('white')
    .setFontSize(8)
    .setWrap(true)
    .setFontFamily('Comfortaa')
    .setVerticalAlignment('middle');
  startRow++;
  
  sheet.getRange(startRow, 1).setValue('% of capacity used');
  sheet.getRange(startRow, 3).setValue(capacityStatus);
  sheet.getRange(startRow, 1, 1, 4).setFontSize(8).setWrap(true).setFontFamily('Comfortaa').setVerticalAlignment('middle');
  
  createProgressBar(sheet, startRow, 4, percentCapacityAllocated);
  
  if (percentCapacityAllocated > 100) {
    sheet.getRange(startRow, 3).setFontColor('#ff0000').setFontWeight('bold');
  }
  startRow++;
  
  sheet.getRange(startRow, 1).setValue('% of Epics with All Stories Pointed');
  sheet.getRange(startRow, 3).setValue(`${percentEpicsWithStoryPoints}%`);
  sheet.getRange(startRow, 1, 1, 4).setFontSize(8).setFontFamily('Comfortaa').setVerticalAlignment('middle');
  
  createProgressBar(sheet, startRow, 4, percentEpicsWithStoryPoints);
  startRow++;
  
  return startRow;
}

function parsePISheetData(values, headers) {
  const issues = [];
  
  // Skip header rows and process data
  for (let i = 4; i < values.length; i++) {
    const row = values[i];
    if (!row[0]) continue; // Skip empty rows
    
    const issue = parsePISheetRow(row, headers);
    issues.push(issue);
  }
  
  return issues;
}

function getScrumTeamsFromPI(piNumber) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const piSheet = spreadsheet.getSheetByName(`PI ${piNumber}`);
  
  if (!piSheet) {
    return [];
  }
  
  const dataRange = piSheet.getDataRange();
  const values = dataRange.getValues();
  const headers = values[3];
  const teamColIndex = headers.indexOf('Scrum Team');
  
  if (teamColIndex === -1) {
    return [];
  }
  
  const scrumTeams = new Set();
  for (let i = 4; i < values.length; i++) {
    const team = values[i][teamColIndex];
    if (team) {
      scrumTeams.add(team);
    }
  }
  
  return Array.from(scrumTeams).sort();
}

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
    
    // If not found, try partial matching
    if (teamRow === -1) {
      const teamNameUpper = teamName.toUpperCase();
      for (let i = 2; i < values.length; i++) {
        const sheetTeamName = values[i][0];
        if (sheetTeamName) {
          const sheetTeamUpper = sheetTeamName.toString().toUpperCase();
          // Check if one contains the other
          if (sheetTeamUpper.includes(teamNameUpper) || teamNameUpper.includes(sheetTeamUpper)) {
            teamRow = i;
            break;
          }
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
    for (let col = 1; col <= 6; col++) {
      total += parseFloat(values[teamRow][col]) || 0;
    }
    
    return { total: total };
    
  } catch (error) {
    console.error('Error reading capacity data:', error);
    return null;
  }
}

function calculateSlottedData(issues, piNumber, scrumTeam) {
  const slottedData = {
    product: { 1: 0, 2: 0, 3: 0, 4: 0, 5: 0, 6: 0, total4: 0, total6: 0 },
    tech: { 1: 0, 2: 0, 3: 0, 4: 0, 5: 0, 6: 0, total4: 0, total6: 0 },
    quality: { 1: 0, 2: 0, 3: 0, 4: 0, 5: 0, 6: 0, total4: 0, total6: 0 }
  };
  
  // Filter for stories AND BUGS belonging to this team (case-insensitive)
  const normalizedScrumTeam = scrumTeam.toUpperCase().replace(/[\s-]/g, '');
  const teamStories = issues.filter(issue => {
    // UPDATED: Include both Story and Bug types
    if ((issue.issueType !== 'Story' && issue.issueType !== 'Bug') || !issue.sprintName) return false;
    
    // Case-insensitive team matching
    const issueTeam = (issue.scrumTeam || '').toUpperCase().replace(/[\s-]/g, '');
    
    return issueTeam === normalizedScrumTeam;
  });
  
  console.log(`Found ${teamStories.length} stories and bugs for team ${scrumTeam} with sprints`);
  
  teamStories.forEach(story => {
    const storyPoints = story.storyPoints || 0;
    if (storyPoints === 0) return;
    
    // Parse sprint name to find iteration - more flexible pattern
    const sprintPattern = new RegExp(`${piNumber}\\s*\\.\\s*(\\d)`, 'i');
    const match = story.sprintName.match(sprintPattern);
    
    if (match) {
      const iteration = parseInt(match[1]);
      if (iteration >= 1 && iteration <= 6) {
        // Determine allocation category using the helper function
        // mapAllocationToCategory and ALLOCATION_CATEGORIES are defined in the main config file
        const category = mapAllocationToCategory(story.allocation);
        
        if (category === ALLOCATION_CATEGORIES.FEATURES) {
          slottedData.product[iteration] += storyPoints;
        } else if (category === ALLOCATION_CATEGORIES.TECH) {
          slottedData.tech[iteration] += storyPoints;
        } else if (category === ALLOCATION_CATEGORIES.QUALITY) {
          // Check if it's planned quality (not unplanned)
          const summary = (story.summary || '').toLowerCase();
          if (!summary.includes('unplanned')) {
            slottedData.quality[iteration] += storyPoints;
          }
        }
      }
    }
  });
  
  // Calculate totals
  for (let i = 1; i <= 5; i++) {
    slottedData.product.total4 += slottedData.product[i];
    slottedData.tech.total4 += slottedData.tech[i];
    slottedData.quality.total4 += slottedData.quality[i];
  }
  
  for (let i = 1; i <= 6; i++) {
    slottedData.product.total6 += slottedData.product[i];
    slottedData.tech.total6 += slottedData.tech[i];
    slottedData.quality.total6 += slottedData.quality[i];
  }
  
  console.log('Slotted data calculated:', slottedData);
  
  return slottedData;
}

/**
 * Get roles and their capacities for a team from the consolidated capacity sheet
 * Combines Before FF and After FF capacities for each role
 * @param {Spreadsheet} spreadsheet - The spreadsheet object
 * @param {string} teamName - Team name to get roles for
 * @param {string} valueStream - Optional value stream filter
 * @returns {Object|null} { roles: { roleName: { beforeFF, afterFF, total, byIteration } }, teamFound: boolean }
 */
function getRolesForTeamFromCapacity(spreadsheet, teamName, valueStream = null) {
  try {
    // Try to find consolidated capacity sheet
    let capacitySheet = null;
    
    if (typeof findCapacityPlanningSheet === 'function') {
      capacitySheet = findCapacityPlanningSheet(spreadsheet);
    }
    
    if (!capacitySheet) {
      // Try common sheet names
      const sheetNames = ['Capacity Planning', 'PI14 - Capacity', 'PI15 - Capacity'];
      for (const name of sheetNames) {
        capacitySheet = spreadsheet.getSheetByName(name);
        if (capacitySheet) break;
      }
      
      // Try pattern matching
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
    
    console.log(`Looking for roles for team "${teamName}" in capacity sheet "${capacitySheet.getName()}"`);
    
    // Find the team in the capacity sheet
    const dataRange = capacitySheet.getDataRange();
    const values = dataRange.getValues();
    const maxRows = values.length;
    const maxCols = values[0] ? values[0].length : 0;
    
    // Normalize team name for matching
    const normalizeTeamName = (name) => name.toUpperCase().replace(/[\s\-_]/g, '');
    const normalizedSearchTeam = normalizeTeamName(teamName);
    
    // Team block configuration
    const teamBlockWidth = 11;
    const teamBlockHeight = 25;
    
    // Find value stream columns (Row 1)
    const valueStreamCols = [];
    for (let col = 0; col < maxCols; col += teamBlockWidth) {
      const vsName = values[0][col];
      if (vsName && vsName.toString().trim()) {
        valueStreamCols.push({ name: vsName.toString().trim(), col: col });
      }
    }
    
    // Search for the team
    let teamLocation = null;
    
    for (const vs of valueStreamCols) {
      // If value stream filter specified, check if it matches
      if (valueStream && !vs.name.toUpperCase().includes(valueStream.toUpperCase())) {
        continue;
      }
      
      // Scan rows for this value stream column
      for (let row = 2; row < maxRows; row += teamBlockHeight) {
        const cellValue = values[row][vs.col];
        if (cellValue && cellValue.toString().trim()) {
          const normalizedCell = normalizeTeamName(cellValue.toString());
          if (normalizedCell === normalizedSearchTeam || 
              normalizedCell.includes(normalizedSearchTeam) ||
              normalizedSearchTeam.includes(normalizedCell)) {
            teamLocation = { row: row, col: vs.col, valueStream: vs.name };
            console.log(`Found team "${teamName}" at row ${row + 1}, column ${vs.col + 1} in ${vs.name}`);
            break;
          }
        }
      }
      if (teamLocation) break;
    }
    
    if (!teamLocation) {
      console.log(`Team "${teamName}" not found in capacity sheet`);
      return null;
    }
    
    // Extract roles from Base Capacity before FF (rows +10 to +16) and after FF (rows +18 to +24)
    const roles = {};
    const totalCol = teamLocation.col + 9; // Total column is at offset +9
    
    // Role rows are at offsets +11 to +15 from team name row (relative to Base Capacity header at +10)
    const beforeFFRoleStartOffset = 11;
    const beforeFFRoleEndOffset = 16;
    const afterFFRoleStartOffset = 19;
    const afterFFRoleEndOffset = 24;
    
    // Read Before FF roles
    for (let offset = beforeFFRoleStartOffset; offset <= beforeFFRoleEndOffset; offset++) {
      const row = teamLocation.row + offset;
      if (row >= maxRows) break;
      
      const roleName = values[row][teamLocation.col];
      const totalValue = values[row][totalCol];
      
      if (roleName && roleName.toString().trim() && 
          !roleName.toString().toLowerCase().includes('base capacity')) {
        const roleKey = roleName.toString().trim().toUpperCase();
        const normalizedRole = ROLE_NORMALIZATION[roleKey] || roleKey;
        
        // Parse total value
        let total = 0;
        if (totalValue && totalValue !== '-' && totalValue !== '') {
          total = parseFloat(totalValue) || 0;
        }
        
        if (!roles[normalizedRole]) {
          roles[normalizedRole] = { 
            displayName: roleName.toString().trim(),
            beforeFF: 0, 
            afterFF: 0, 
            total: 0,
            byIteration: { 1: 0, 2: 0, 3: 0, 4: 0, 5: 0, 6: 0 }
          };
        }
        roles[normalizedRole].beforeFF = total;
        
        // Read per-iteration values (columns +3 to +8 are iterations 1-6)
        for (let iter = 1; iter <= 6; iter++) {
          const iterCol = teamLocation.col + 2 + iter; // +3 for iter 1, +4 for iter 2, etc.
          const iterValue = values[row][iterCol];
          if (iterValue && iterValue !== '-' && iterValue !== '') {
            roles[normalizedRole].byIteration[iter] = parseFloat(iterValue) || 0;
          }
        }
      }
    }
    
    // Read After FF roles and add to totals
    for (let offset = afterFFRoleStartOffset; offset <= afterFFRoleEndOffset; offset++) {
      const row = teamLocation.row + offset;
      if (row >= maxRows) break;
      
      const roleName = values[row][teamLocation.col];
      const totalValue = values[row][totalCol];
      
      if (roleName && roleName.toString().trim() && 
          !roleName.toString().toLowerCase().includes('base capacity')) {
        const roleKey = roleName.toString().trim().toUpperCase();
        const normalizedRole = ROLE_NORMALIZATION[roleKey] || roleKey;
        
        let total = 0;
        if (totalValue && totalValue !== '-' && totalValue !== '') {
          total = parseFloat(totalValue) || 0;
        }
        
        if (!roles[normalizedRole]) {
          roles[normalizedRole] = { 
            displayName: roleName.toString().trim(),
            beforeFF: 0, 
            afterFF: 0, 
            total: 0,
            byIteration: { 1: 0, 2: 0, 3: 0, 4: 0, 5: 0, 6: 0 }
          };
        }
        roles[normalizedRole].afterFF = total;
      }
    }
    
    // Calculate combined totals
    Object.keys(roles).forEach(roleKey => {
      roles[roleKey].total = roles[roleKey].beforeFF + roles[roleKey].afterFF;
    });
    
    // Filter out roles with zero capacity
    const activeRoles = {};
    Object.keys(roles).forEach(roleKey => {
      if (roles[roleKey].total > 0) {
        activeRoles[roleKey] = roles[roleKey];
      }
    });
    
    console.log(`Found ${Object.keys(activeRoles).length} active roles for team "${teamName}":`, 
      Object.keys(activeRoles).map(r => `${r}: ${activeRoles[r].total}`).join(', '));
    
    return { roles: activeRoles, teamFound: true, valueStream: teamLocation.valueStream };
    
  } catch (error) {
    console.error(`Error getting roles for team ${teamName}:`, error);
    return null;
  }
}

/**
 * Detect role from a ticket's labels or title
 * @param {Object} issue - The issue object with labels and summary
 * @returns {string|null} Normalized role name or null if not detected
 */
function detectRoleFromTicket(issue) {
  const labels = issue.labels || [];
  const summary = issue.summary || '';
  
  // Define role patterns for detection
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
  
  // Method 1: Check labels first (higher priority)
  for (const label of labels) {
    const labelUpper = label.toString().toUpperCase().trim();
    const normalizedRole = ROLE_NORMALIZATION[labelUpper];
    if (normalizedRole) {
      return normalizedRole;
    }
    
    // Also check patterns against labels
    for (const rp of rolePatterns) {
      if (rp.pattern.test(label)) {
        return rp.role;
      }
    }
  }
  
  // Method 2: Check title prefix patterns
  // Common patterns: [BE], (BE), BE:, BE -, BE at start
  const prefixPatterns = [
    /^\s*\[([A-Z\-]+)\]/i,      // [BE] at start
    /^\s*\(([A-Z\-]+)\)/i,      // (BE) at start
    /^\s*([A-Z\-]+)\s*:/i,      // BE: at start
    /^\s*([A-Z\-]+)\s*-\s/i,    // BE - at start
    /^\s*([A-Z\-]{2,6})\s+/i    // BE word at start (2-6 chars)
  ];
  
  for (const pattern of prefixPatterns) {
    const match = summary.match(pattern);
    if (match) {
      const extracted = match[1].toUpperCase().trim();
      const normalizedRole = ROLE_NORMALIZATION[extracted];
      if (normalizedRole) {
        return normalizedRole;
      }
      
      // Check role patterns
      for (const rp of rolePatterns) {
        if (rp.pattern.test(extracted)) {
          return rp.role;
        }
      }
    }
  }
  
  // Method 3: Check if summary contains role keywords (less reliable, only for clear cases)
  // Only match if it appears to be a role indicator, not just part of text
  const roleIndicatorPatterns = [
    /\[([A-Z\-]+)\]/i,          // [BE] anywhere
    /\(([A-Z\-]+)\)/i           // (BE) anywhere
  ];
  
  for (const pattern of roleIndicatorPatterns) {
    const match = summary.match(pattern);
    if (match) {
      const extracted = match[1].toUpperCase().trim();
      const normalizedRole = ROLE_NORMALIZATION[extracted];
      if (normalizedRole) {
        return normalizedRole;
      }
    }
  }
  
  return null;
}

/**
 * Calculate role-based slotted data from issues
 * @param {Array} issues - All issues
 * @param {number} piNumber - PI number
 * @param {string} scrumTeam - Scrum team name
 * @param {Object} availableRoles - Roles available for this team from capacity sheet
 * @returns {Object} Role data by iteration
 */
function calculateRoleSlottedData(issues, piNumber, scrumTeam, availableRoles) {
  const roleData = {};
  
  // Initialize role data structure
  const roleKeys = Object.keys(availableRoles || {});
  roleKeys.forEach(role => {
    roleData[role] = { 
      1: 0, 2: 0, 3: 0, 4: 0, 5: 0, 6: 0, 
      total: 0,
      capacity: availableRoles[role].total || 0,
      displayName: availableRoles[role].displayName || role
    };
  });
  
  // Add Unassigned category
  roleData['Unassigned'] = { 
    1: 0, 2: 0, 3: 0, 4: 0, 5: 0, 6: 0, 
    total: 0,
    capacity: 0,
    displayName: 'Unassigned'
  };
  
  // Filter for stories and bugs belonging to this team
  const normalizedScrumTeam = scrumTeam.toUpperCase().replace(/[\s-]/g, '');
  const teamStories = issues.filter(issue => {
    if ((issue.issueType !== 'Story' && issue.issueType !== 'Bug') || !issue.sprintName) return false;
    const issueTeam = (issue.scrumTeam || '').toUpperCase().replace(/[\s-]/g, '');
    return issueTeam === normalizedScrumTeam;
  });
  
  console.log(`Processing ${teamStories.length} stories/bugs for role breakdown`);
  
  teamStories.forEach(story => {
    const storyPoints = story.storyPoints || 0;
    if (storyPoints === 0) return;
    
    // Parse sprint name to find iteration
    const sprintPattern = new RegExp(`${piNumber}\\s*\\.\\s*(\\d)`, 'i');
    const match = story.sprintName.match(sprintPattern);
    
    if (match) {
      const iteration = parseInt(match[1]);
      if (iteration >= 1 && iteration <= 6) {
        // Detect role from ticket
        let detectedRole = detectRoleFromTicket(story);
        
        // If detected role exists in available roles, use it; otherwise check normalization
        if (detectedRole && !roleData[detectedRole]) {
          // Role detected but not in capacity - might be a variation
          const normalized = ROLE_NORMALIZATION[detectedRole];
          if (normalized && roleData[normalized]) {
            detectedRole = normalized;
          } else {
            detectedRole = null; // Fall back to Unassigned
          }
        }
        
        const roleKey = detectedRole || 'Unassigned';
        
        if (roleData[roleKey]) {
          roleData[roleKey][iteration] += storyPoints;
          roleData[roleKey].total += storyPoints;
        }
      }
    }
  });
  
  console.log('Role breakdown calculated:', 
    Object.keys(roleData).map(r => `${r}: ${roleData[r].total}`).join(', '));
  
  return roleData;
}

/**
 * Create Role Breakdown chart on the sheet
 * @param {Sheet} sheet - The sheet to write to
 * @param {number} startRow - Starting row number
 * @param {Array} issues - All issues
 * @param {string} scrumTeam - Scrum team name
 * @param {string} programIncrement - PI string (e.g., "PI 14")
 * @param {Spreadsheet} spreadsheet - The spreadsheet object
 * @returns {number} Next available row after the chart
 */
function createRoleBreakdownChart(sheet, startRow, issues, scrumTeam, programIncrement, spreadsheet) {
  console.log(`Creating Role Breakdown chart for ${scrumTeam}`);
  
  // Extract PI number
  const piNumber = parseInt(programIncrement.replace('PI ', ''));
  if (isNaN(piNumber)) {
    console.error('Invalid PI number in programIncrement:', programIncrement);
    return startRow;
  }
  
  // Get roles for this team from capacity sheet
  const roleCapacity = getRolesForTeamFromCapacity(spreadsheet, scrumTeam);
  
  if (!roleCapacity || !roleCapacity.roles || Object.keys(roleCapacity.roles).length === 0) {
    console.log(`No roles found for team ${scrumTeam} - skipping Role Breakdown section`);
    return startRow;
  }
  
  // Calculate role-based slotted data
  const roleData = calculateRoleSlottedData(issues, piNumber, scrumTeam, roleCapacity.roles);
  
  // Check if there's any data to show (at least one role with capacity or usage)
  const hasData = Object.keys(roleData).some(role => 
    roleData[role].total > 0 || roleData[role].capacity > 0
  );
  
  if (!hasData) {
    console.log(`No role data found for team ${scrumTeam} - skipping Role Breakdown section`);
    return startRow;
  }
  
  // Filter roles to show (has capacity or has usage)
  const rolesToShow = Object.keys(roleData).filter(role => 
    roleData[role].capacity > 0 || roleData[role].total > 0
  );
  
  // Sort roles: named roles first (by capacity descending), Unassigned last
  rolesToShow.sort((a, b) => {
    if (a === 'Unassigned') return 1;
    if (b === 'Unassigned') return -1;
    return (roleData[b].capacity || 0) - (roleData[a].capacity || 0);
  });
  
  // === RENDER THE CHART ===
  
  // Title row
  sheet.getRange(startRow, 1).setValue('Role Breakdown');
  sheet.getRange(startRow, 1, 1, 12).setBackground('#E1D5E7');
  sheet.getRange(startRow, 1).setFontSize(14).setFontWeight('bold').setFontColor('black');
  sheet.getRange(startRow, 1).setFontFamily('Comfortaa');
  sheet.getRange(startRow, 1).setVerticalAlignment('middle');
  setRowHeightWithLimit(sheet, startRow, 30, 70);
  startRow++;
  
  // Headers
  const headers = [
    'Role', 'Baseline Capacity',
    `${piNumber}.1`, `${piNumber}.2`, `${piNumber}.3`, `${piNumber}.4`, `${piNumber}.5`, `${piNumber}.6`,
    'Total Used', 'Remaining for Use', '% Remaining'
  ];
  
  sheet.getRange(startRow, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(startRow, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#9b7bb8')
    .setFontColor('white')
    .setFontSize(8)
    .setWrap(true)
    .setFontFamily('Comfortaa')
    .setVerticalAlignment('middle')
    .setHorizontalAlignment('center');
  
  setRowHeightWithLimit(sheet, startRow, 30, 70);
  
  const headerRow = startRow;
  startRow++;
  const dataStartRow = startRow;
  
  // Data rows
  rolesToShow.forEach((roleKey, index) => {
    const role = roleData[roleKey];
    const capacity = Math.ceil(role.capacity || 0);
    const totalUsed = Math.ceil(role.total || 0);
    const remaining = capacity - totalUsed;
    const percentRemaining = capacity > 0 ? Math.round((remaining / capacity) * 100) : (totalUsed > 0 ? -100 : 0);
    
    const rowData = [
      role.displayName || roleKey,
      capacity,
      Math.round(role[1] || 0),
      Math.round(role[2] || 0),
      Math.round(role[3] || 0),
      Math.round(role[4] || 0),
      Math.round(role[5] || 0),
      Math.round(role[6] || 0),
      totalUsed,
      remaining,
      percentRemaining + '%'
    ];
    
    sheet.getRange(startRow, 1, 1, rowData.length).setValues([rowData]);
    
    // Formatting
    sheet.getRange(startRow, 1, 1, headers.length)
      .setFontSize(8)
      .setWrap(true)
      .setFontFamily('Comfortaa')
      .setVerticalAlignment('middle');
    
    sheet.getRange(startRow, 2, 1, headers.length - 1).setHorizontalAlignment('center');
    
    // Alternate row coloring
    if (index % 2 === 1) {
      sheet.getRange(startRow, 1, 1, headers.length).setBackground('#f5f5f5');
    }
    
    // Color code the Remaining column (column 10)
    if (remaining >= 0) {
      sheet.getRange(startRow, 10).setBackground('#ccffcc'); // Green for positive
    } else {
      sheet.getRange(startRow, 10).setBackground('#ffcccc'); // Red for negative
    }
    
    // Color code % Remaining (column 11)
    if (percentRemaining >= 0) {
      sheet.getRange(startRow, 11).setBackground('#ccffcc');
    } else {
      sheet.getRange(startRow, 11).setBackground('#ffcccc');
    }
    
    // Highlight Unassigned row differently
    if (roleKey === 'Unassigned' && role.total > 0) {
      sheet.getRange(startRow, 1).setBackground('#fff3cd'); // Light yellow warning
      sheet.getRange(startRow, 9).setBackground('#fff3cd');
    }
    
    setRowHeightWithLimit(sheet, startRow, 25, 70);
    startRow++;
  });
  
  const dataEndRow = startRow - 1;
  
  // Add totals row
  const totalCapacity = rolesToShow.reduce((sum, r) => sum + (roleData[r].capacity || 0), 0);
  const totalUsed = rolesToShow.reduce((sum, r) => sum + (roleData[r].total || 0), 0);
  const totalRemaining = totalCapacity - totalUsed;
  const totalPercentRemaining = totalCapacity > 0 ? Math.round((totalRemaining / totalCapacity) * 100) : 0;
  
  // Sum iterations
  const iterTotals = [0, 0, 0, 0, 0, 0];
  rolesToShow.forEach(r => {
    for (let i = 1; i <= 6; i++) {
      iterTotals[i - 1] += roleData[r][i] || 0;
    }
  });
  
  const totalRowData = [
    'TOTAL',
    Math.ceil(totalCapacity),
    Math.round(iterTotals[0]),
    Math.round(iterTotals[1]),
    Math.round(iterTotals[2]),
    Math.round(iterTotals[3]),
    Math.round(iterTotals[4]),
    Math.round(iterTotals[5]),
    Math.ceil(totalUsed),
    Math.ceil(totalRemaining),
    totalPercentRemaining + '%'
  ];
  
  sheet.getRange(startRow, 1, 1, totalRowData.length).setValues([totalRowData]);
  sheet.getRange(startRow, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#e0e0e0')
    .setFontSize(8)
    .setWrap(true)
    .setFontFamily('Comfortaa')
    .setVerticalAlignment('middle');
  sheet.getRange(startRow, 2, 1, headers.length - 1).setHorizontalAlignment('center');
  
  // Color totals remaining
  if (totalRemaining >= 0) {
    sheet.getRange(startRow, 10).setBackground('#ccffcc');
  } else {
    sheet.getRange(startRow, 10).setBackground('#ffcccc');
  }
  if (totalPercentRemaining >= 0) {
    sheet.getRange(startRow, 11).setBackground('#ccffcc');
  } else {
    sheet.getRange(startRow, 11).setBackground('#ffcccc');
  }
  
  setRowHeightWithLimit(sheet, startRow, 25, 70);
  
  // Add borders around the table
  sheet.getRange(headerRow, 1, startRow - headerRow + 1, headers.length).setBorder(
    true, true, true, true, false, false,
    '#555555', SpreadsheetApp.BorderStyle.SOLID_MEDIUM
  );
  
  // Add thick border between header and data
  sheet.getRange(headerRow, 1, 1, headers.length).setBorder(
    true, true, true, true, false, false,
    '#333333', SpreadsheetApp.BorderStyle.SOLID_THICK
  );
  
  // Add thick border above totals row
  sheet.getRange(startRow, 1, 1, headers.length).setBorder(
    true, true, true, true, false, false,
    '#333333', SpreadsheetApp.BorderStyle.SOLID_THICK
  );
  
  startRow++;
  
  console.log(`Role Breakdown chart created with ${rolesToShow.length} roles`);
  
  return startRow + 1;
}

// ===== TEAM INITIATIVE ANALYSIS FUNCTIONS =====

/**
 * Create Initiative Analysis section for team summary
 * Shows Portfolio Initiative and Program Initiative distribution with pie charts
 * @param {Sheet} sheet - The sheet to write to
 * @param {number} startRow - Starting row
 * @param {Array} issues - All issues for this team
 * @param {string} scrumTeam - Team name
 * @param {string} programIncrement - PI string (e.g., "PI 14")
 * @returns {number} Next available row
 */
function createTeamInitiativeAnalysis(sheet, startRow, issues, scrumTeam, programIncrement) {
  console.log(`Creating Initiative Analysis for ${scrumTeam}`);
  
  // Filter for epics only
  const epics = issues.filter(issue => issue.issueType === 'Epic');
  
  if (epics.length === 0) {
    console.log(`No epics found for ${scrumTeam} - skipping Initiative Analysis`);
    return startRow;
  }
  
  // Check if any epics have initiative data
  const hasPortfolioData = epics.some(e => e.portfolioInitiative && e.portfolioInitiative.trim());
  const hasProgramData = epics.some(e => e.programInitiative && e.programInitiative.trim());
  
  if (!hasPortfolioData && !hasProgramData) {
    console.log(`No initiative data found for ${scrumTeam} - skipping Initiative Analysis`);
    return startRow;
  }
  
  let currentRow = startRow;
  
  // Section title
  sheet.getRange(currentRow, 1).setValue('Initiative Analysis');
  sheet.getRange(currentRow, 1, 1, 12).setBackground('#E1D5E7');
  sheet.getRange(currentRow, 1).setFontSize(14).setFontWeight('bold').setFontColor('black');
  sheet.getRange(currentRow, 1).setFontFamily('Comfortaa');
  sheet.getRange(currentRow, 1).setVerticalAlignment('middle');
  setRowHeightWithLimit(sheet, currentRow, 30, 70);
  currentRow += 2;
  
  // ===== PORTFOLIO INITIATIVE SECTION =====
  if (hasPortfolioData) {
    currentRow = writeTeamInitiativeTable(
      sheet, 
      currentRow, 
      epics, 
      'portfolioInitiative', 
      'Portfolio Initiative',
      true  // Include pie chart
    );
    currentRow += 2;
  }
  
  // ===== PROGRAM INITIATIVE SECTION =====
  if (hasProgramData) {
    currentRow = writeTeamInitiativeTable(
      sheet, 
      currentRow, 
      epics, 
      'programInitiative', 
      'Program Initiative',
      true  // Include pie chart
    );
    currentRow += 2;
  }
  
  console.log(`Initiative Analysis created for ${scrumTeam}`);
  return currentRow;
}

/**
 * Write initiative distribution table with optional pie chart
 * @param {Sheet} sheet - The sheet to write to
 * @param {number} startRow - Starting row
 * @param {Array} epics - Epic issues
 * @param {string} field - Field name ('portfolioInitiative' or 'programInitiative')
 * @param {string} label - Display label
 * @param {boolean} includeChart - Whether to include a pie chart
 * @returns {number} Next available row
 */
function writeTeamInitiativeTable(sheet, startRow, epics, field, label, includeChart = true) {
  let currentRow = startRow;
  
  // Sub-section header
  sheet.getRange(currentRow, 1).setValue(`${label} Distribution`);
  sheet.getRange(currentRow, 1, 1, 5)
    .setFontSize(11)
    .setFontWeight('bold')
    .setBackground('#9b7bb8')
    .setFontColor('white')
    .setFontFamily('Comfortaa');
  setRowHeightWithLimit(sheet, currentRow, 25, 70);
  currentRow++;
  
  // Calculate distribution
  const distribution = calculateTeamInitiativeDistribution(epics, field);
  
  if (distribution.rows.length === 0) {
    sheet.getRange(currentRow, 1).setValue(`No ${label} data available`);
    sheet.getRange(currentRow, 1).setFontStyle('italic').setFontColor('#666666').setFontFamily('Comfortaa');
    return currentRow + 1;
  }
  
  // Headers
  const headers = [label, 'Epic Count', 'Total Points', '% of Total'];
  sheet.getRange(currentRow, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(currentRow, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#1B365D')
    .setFontColor('white')
    .setFontSize(8)
    .setFontFamily('Comfortaa')
    .setHorizontalAlignment('center');
  setRowHeightWithLimit(sheet, currentRow, 25, 70);
  
  const headerRow = currentRow;
  currentRow++;
  const dataStartRow = currentRow;
  
  // Data rows
  sheet.getRange(currentRow, 1, distribution.rows.length, headers.length)
    .setValues(distribution.rows);
  sheet.getRange(currentRow, 1, distribution.rows.length, headers.length)
    .setFontSize(8)
    .setFontFamily('Comfortaa')
    .setVerticalAlignment('middle');
  sheet.getRange(currentRow, 2, distribution.rows.length, 3)
    .setHorizontalAlignment('center');
  
  // Alternate row coloring
  for (let i = 0; i < distribution.rows.length; i++) {
    if (i % 2 === 1) {
      sheet.getRange(currentRow + i, 1, 1, headers.length).setBackground('#f5f5f5');
    }
    setRowHeightWithLimit(sheet, currentRow + i, 22, 70);
  }
  
  currentRow += distribution.rows.length;
  
  // Totals row
  const totals = ['TOTAL', distribution.totalEpics, Math.ceil(distribution.totalPoints), '100%'];
  sheet.getRange(currentRow, 1, 1, headers.length).setValues([totals]);
  sheet.getRange(currentRow, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#FFC72C')
    .setFontSize(8)
    .setFontFamily('Comfortaa');
  sheet.getRange(currentRow, 2, 1, 3).setHorizontalAlignment('center');
  setRowHeightWithLimit(sheet, currentRow, 25, 70);
  
  // Add borders
  sheet.getRange(headerRow, 1, distribution.rows.length + 2, headers.length).setBorder(
    true, true, true, true, false, false,
    '#555555', SpreadsheetApp.BorderStyle.SOLID_MEDIUM
  );
  
  currentRow++;
  
  // Create pie chart if requested and we have data
  if (includeChart && distribution.rows.length > 0) {
    try {
      createTeamInitiativePieChart(
        sheet,
        dataStartRow,
        Math.min(distribution.rows.length, 10),  // Limit to top 10 for readability
        label,
        6,  // Chart column (F)
        dataStartRow - 2  // Position chart at section header row
      );
    } catch (e) {
      console.error(`Error creating ${label} pie chart:`, e);
    }
  }
  
  return currentRow;
}

/**
 * Calculate initiative distribution for a team
 * @param {Array} epics - Epic issues
 * @param {string} field - Field name to group by
 * @returns {Object} { rows: [[name, count, points, %], ...], totalEpics, totalPoints }
 */
function calculateTeamInitiativeDistribution(epics, field) {
  const distribution = {};
  let totalPoints = 0;
  
  epics.forEach(epic => {
    const initiative = epic[field] || 'Not Specified';
    const points = calculateTeamEpicPoints(epic);
    
    if (!distribution[initiative]) {
      distribution[initiative] = { epicCount: 0, points: 0 };
    }
    distribution[initiative].epicCount++;
    distribution[initiative].points += points;
    totalPoints += points;
  });
  
  // Convert to sorted array (descending by points), truncate long names
  const maxLen = 50;
  const sorted = Object.entries(distribution)
    .sort((a, b) => b[1].points - a[1].points)
    .map(([name, data]) => [
      name.length > maxLen ? name.substring(0, maxLen - 3) + '...' : name,
      data.epicCount,
      Math.ceil(data.points),
      totalPoints > 0 ? Math.round((data.points / totalPoints) * 100) + '%' : '0%'
    ]);
  
  return {
    rows: sorted,
    totalEpics: epics.length,
    totalPoints: totalPoints
  };
}

/**
 * Calculate points for an epic (Feature Points x 10 or Story Point Estimate)
 * @param {Object} epic - Epic object
 * @returns {number} Calculated points
 */
function calculateTeamEpicPoints(epic) {
  // Use Feature Points x 10 if available
  if (epic.featurePoints && epic.featurePoints > 0) {
    return epic.featurePoints * 10;
  }
  // Fall back to Story Point Estimate
  if (epic.storyPointEstimate && epic.storyPointEstimate > 0) {
    return epic.storyPointEstimate;
  }
  // Fall back to aggregated story points
  return epic.storyPoints || 0;
}

/**
 * Create pie chart for team initiative distribution
 * @param {Sheet} sheet - Target sheet
 * @param {number} dataStartRow - First row of data
 * @param {number} dataRowCount - Number of data rows
 * @param {string} title - Chart title
 * @param {number} chartColumn - Column to position chart
 * @param {number} chartRow - Row to position chart
 */
function createTeamInitiativePieChart(sheet, dataStartRow, dataRowCount, title, chartColumn, chartRow) {
  try {
    const labelRange = sheet.getRange(dataStartRow, 1, dataRowCount, 1);  // Initiative names
    const valueRange = sheet.getRange(dataStartRow, 3, dataRowCount, 1);  // Points column
    
    const chart = sheet.newChart()
      .setChartType(Charts.ChartType.PIE)
      .addRange(labelRange)
      .addRange(valueRange)
      .setPosition(chartRow, chartColumn, 0, 0)
      .setOption('title', title + ' by Points')
      .setOption('width', 400)
      .setOption('height', 280)
      .setOption('pieSliceText', 'percentage')
      .setOption('legend', { position: 'right', textStyle: { fontSize: 8 } })
      .setOption('titleTextStyle', { fontSize: 10, bold: true })
      .setOption('colors', [
        '#1B365D', '#6B3FA0', '#FFC72C', '#4285F4', '#34A853',
        '#EA4335', '#FBBC05', '#9AA0A6', '#5F6368', '#F28B82'
      ])
      .build();
    
    sheet.insertChart(chart);
    console.log(`Created pie chart: ${title}`);
    
  } catch (error) {
    console.error(`Error creating pie chart "${title}":`, error);
  }
}

function createIterationSlottingChart(sheet, startRow, issues, scrumTeam, programIncrement, spreadsheet) {
  console.log(`Creating Iteration Slotting chart for ${scrumTeam}`);
  
  // Extract PI number
  const piNumber = parseInt(programIncrement.replace('PI ', ''));
  if (isNaN(piNumber)) {
    console.error('Invalid PI number in programIncrement:', programIncrement);
    return startRow;
  }
  
  // Get spreadsheet reference if not passed
  if (!spreadsheet) {
    spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  }
  
  // Try to get capacity data from consolidated capacity sheet
  let capacityData = null;
  try {
    if (typeof getCapacityDataForTeamConsolidated === 'function') {
      capacityData = getCapacityDataForTeamConsolidated(spreadsheet, scrumTeam);
      if (capacityData) {
        console.log(`Found capacity data for ${scrumTeam} in consolidated format`);
      }
    }
  } catch (e) {
    console.log(`Error getting consolidated capacity data: ${e.message}`);
  }
  
  // If no capacity data found, team won't have iteration slotting
  if (!capacityData) {
    console.log(`No capacity data found for team ${scrumTeam} - skipping iteration slotting`);
    return startRow;
  }
  
  console.log(`Using capacity data for team ${scrumTeam}`);
  
  // Title - fill columns A through L with purple
  sheet.getRange(startRow, 1).setValue('Iteration Slotting');
  sheet.getRange(startRow, 1, 1, 12).setBackground('#E1D5E7');
  sheet.getRange(startRow, 1).setFontSize(14).setFontWeight('bold').setFontColor('black');
  sheet.getRange(startRow, 1).setFontFamily('Comfortaa');
  sheet.getRange(startRow, 1).setVerticalAlignment('middle');
  setRowHeightWithLimit(sheet, startRow, 30, 70);
  
  const headerRow = startRow + 1;
  
  // Headers
  const headers = [
    'Iteration', 'Baseline Capacity', 'Product Load', 'Slotted Product Load', 'Remaining for Use',
    'Tech/Platform Load', 'Slotted Tech/Platform Load', 'Remaining for Use',
    'Planned Quality Load', 'Slotted Planned Quality Load', 'Remaining for Use',
    'Unplanned Work'
  ];
  
  sheet.getRange(headerRow, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(headerRow, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#9b7bb8')
    .setFontColor('white')
    .setFontSize(8)
    .setWrap(true)
    .setFontFamily('Comfortaa')
    .setVerticalAlignment('middle');
  
  setRowHeightWithLimit(sheet, headerRow, 50, 70);
  
  // Set column widths
  for (let col = 1; col <= headers.length; col++) {
    sheet.setColumnWidth(col, 100);
  }
  
  // Calculate slotted values from PI sheet data
  const slottedData = calculateSlottedData(issues, piNumber, scrumTeam);
  
  // Data rows
  const iterations = [
    'Iteration 1', 'Iteration 2', 'Iteration 3', 'Iteration 4', 
    'Iteration 5', 'Iteration 6', 'Total (5 iterations)', 'Total (6 iterations)'
  ];
  
  const dataStartRow = headerRow + 1;
  
  iterations.forEach((iteration, index) => {
    const currentRow = dataStartRow + index;
    const iterationNum = index + 1;
    
    // Alternate row colors for regular iterations
    if (iterationNum <= 6) {
      if (iterationNum % 2 === 0) {
        sheet.getRange(currentRow, 1, 1, headers.length).setBackground('#f5f5f5');
      }
    }
    
    // Iteration column
    sheet.getRange(currentRow, 1).setValue(iteration);
    sheet.getRange(currentRow, 1).setWrap(true);
    
    if (iterationNum <= 6) {
      // Regular iterations - use values from capacityData
      const iterData = capacityData.byIteration ? capacityData.byIteration[iterationNum] : null;
      
      // Baseline Capacity - use base capacity before FF, distributed by iteration
      // Note: The consolidated format stores total base capacity, so we estimate per iteration
      const baseCapPerIter = Math.round((capacityData.baseCapacityBeforeFF || 0) / 6);
      sheet.getRange(currentRow, 2).setValue(baseCapPerIter);
      
      if (iterData) {
        // Product Load (Feature + Compliance)
        const productLoad = Math.round((iterData.productFeature || 0) + (iterData.productCompliance || 0));
        sheet.getRange(currentRow, 3).setValue(productLoad);
        
        // Tech/Platform Load
        sheet.getRange(currentRow, 6).setValue(Math.round(iterData.techPlatform || 0));
        
        // Planned Quality Load
        sheet.getRange(currentRow, 9).setValue(Math.round(iterData.quality || 0));
        
        // Unplanned Work (KLO)
        sheet.getRange(currentRow, 12).setValue(Math.round(iterData.klo || 0));
      } else {
        // Fallback to allocation totals distributed evenly if no iteration data
        const productLoad = Math.round((capacityData.productCapacity || 0) / 6);
        sheet.getRange(currentRow, 3).setValue(productLoad);
        sheet.getRange(currentRow, 6).setValue(Math.round((capacityData.allocations?.techPlatform || 0) / 6));
        sheet.getRange(currentRow, 9).setValue(Math.round((capacityData.allocations?.quality || 0) / 6));
        sheet.getRange(currentRow, 12).setValue(Math.round((capacityData.allocations?.klo || 0) / 6));
      }
      
      // Slotted Product Load
      sheet.getRange(currentRow, 4).setValue(Math.round(slottedData.product[iterationNum] || 0));
      
      // Remaining (Product)
      sheet.getRange(currentRow, 5).setFormula(`=C${currentRow}-D${currentRow}`);
      
      // Slotted Tech/Platform Load
      sheet.getRange(currentRow, 7).setValue(Math.round(slottedData.tech[iterationNum] || 0));
      
      // Remaining (Tech)
      sheet.getRange(currentRow, 8).setFormula(`=F${currentRow}-G${currentRow}`);
      
      // Slotted Planned Quality Load
      sheet.getRange(currentRow, 10).setValue(Math.round(slottedData.quality[iterationNum] || 0));
      
      // Remaining (Quality)
      sheet.getRange(currentRow, 11).setFormula(`=I${currentRow}-J${currentRow}`);
      
      // Unplanned Work column always has light grey background
      sheet.getRange(currentRow, 12).setBackground('#f5f5f5');
      
    } else if (iterationNum === 7) {
      // Total (5 iterations)
      sheet.getRange(currentRow, 2).setFormula(`=SUM(B${dataStartRow}:B${dataStartRow + 4})`);
      sheet.getRange(currentRow, 3).setFormula(`=SUM(C${dataStartRow}:C${dataStartRow + 4})`);
      sheet.getRange(currentRow, 4).setValue(Math.round(slottedData.product.total4 || 0));
      sheet.getRange(currentRow, 5).setFormula(`=C${currentRow}-D${currentRow}`);
      sheet.getRange(currentRow, 6).setFormula(`=SUM(F${dataStartRow}:F${dataStartRow + 4})`);
      sheet.getRange(currentRow, 7).setValue(Math.round(slottedData.tech.total4 || 0));
      sheet.getRange(currentRow, 8).setFormula(`=F${currentRow}-G${currentRow}`);
      sheet.getRange(currentRow, 9).setFormula(`=SUM(I${dataStartRow}:I${dataStartRow + 4})`);
      sheet.getRange(currentRow, 10).setValue(Math.round(slottedData.quality.total4 || 0)); 
      sheet.getRange(currentRow, 11).setFormula(`=I${currentRow}-J${currentRow}`);
      sheet.getRange(currentRow, 12).setFormula(`=SUM(L${dataStartRow}:L${dataStartRow + 4})`);
      
      // Bold and darker grey background for totals
      sheet.getRange(currentRow, 1, 1, headers.length).setFontWeight('bold');
      sheet.getRange(currentRow, 1, 1, headers.length).setBackground('#e0e0e0');
      sheet.getRange(currentRow, 12).setBackground('#f5f5f5');
      
    } else if (iterationNum === 8) {
      // Total (6 iterations)
      sheet.getRange(currentRow, 2).setFormula(`=SUM(B${dataStartRow}:B${dataStartRow + 5})`);
      sheet.getRange(currentRow, 3).setFormula(`=SUM(C${dataStartRow}:C${dataStartRow + 5})`);
      sheet.getRange(currentRow, 4).setValue(Math.round(slottedData.product.total6 || 0));
      sheet.getRange(currentRow, 5).setFormula(`=C${currentRow}-D${currentRow}`);
      sheet.getRange(currentRow, 6).setFormula(`=SUM(F${dataStartRow}:F${dataStartRow + 5})`);
      sheet.getRange(currentRow, 7).setValue(Math.round(slottedData.tech.total6 || 0));
      sheet.getRange(currentRow, 8).setFormula(`=F${currentRow}-G${currentRow}`);
      sheet.getRange(currentRow, 9).setFormula(`=SUM(I${dataStartRow}:I${dataStartRow + 5})`);
      sheet.getRange(currentRow, 10).setValue(Math.round(slottedData.quality.total6 || 0));
      sheet.getRange(currentRow, 11).setFormula(`=I${currentRow}-J${currentRow}`);
      sheet.getRange(currentRow, 12).setFormula(`=SUM(L${dataStartRow}:L${dataStartRow + 5})`);
      
      // Bold and darker grey background for totals
      sheet.getRange(currentRow, 1, 1, headers.length).setFontWeight('bold');
      sheet.getRange(currentRow, 1, 1, headers.length).setBackground('#e0e0e0');
      sheet.getRange(currentRow, 12).setBackground('#f5f5f5');
    }
  });
  
  // Add thick border after iteration 4 for PI 12, after iteration 5 for others
  const borderAfterRow = piNumber === 12 ? dataStartRow + 3 : dataStartRow + 4;
  sheet.getRange(borderAfterRow + 1, 1, 1, headers.length).setBorder(
    true, false, false, false, false, false, 
    '#000000', SpreadsheetApp.BorderStyle.SOLID_THICK
  );
  
  // Add border around totals rows
  sheet.getRange(dataStartRow + 6, 1, 2, headers.length).setBorder(
    true, true, true, true, false, false,
    '#666666', SpreadsheetApp.BorderStyle.SOLID_MEDIUM
  );
  
  // Format data area
  sheet.getRange(dataStartRow, 1, iterations.length, headers.length).setFontSize(8).setWrap(true).setFontFamily('Comfortaa').setVerticalAlignment('middle');
  sheet.getRange(dataStartRow, 2, iterations.length, headers.length - 1).setHorizontalAlignment('center');
  
  // Set standard row heights
  for (let i = 0; i < iterations.length; i++) {
    setRowHeightWithLimit(sheet, dataStartRow + i, 25, 70);
  }
  
  // Add dark grey borders around specific column groups
  // C10:E18 (Product columns)
  sheet.getRange(dataStartRow - 1, 3, 9, 3).setBorder(
    true, true, true, true, false, false,
    '#555555', SpreadsheetApp.BorderStyle.SOLID_THICK
  );
  
  // F10:H18 (Tech columns)
  sheet.getRange(dataStartRow - 1, 6, 9, 3).setBorder(
    true, true, true, true, false, false,
    '#555555', SpreadsheetApp.BorderStyle.SOLID_THICK
  );
  
  // I10:K18 (Quality columns)
  sheet.getRange(dataStartRow - 1, 9, 9, 3).setBorder(
    true, true, true, true, false, false,
    '#555555', SpreadsheetApp.BorderStyle.SOLID_THICK
  );
  
  // L10:L18 (Unplanned Work column)
  sheet.getRange(dataStartRow - 1, 12, 9, 1).setBorder(
    true, true, true, true, false, false,
    '#555555', SpreadsheetApp.BorderStyle.SOLID_THICK
  );
  
  // Apply conditional formatting to Remaining columns AFTER row coloring
  SpreadsheetApp.flush();
  
  for (let i = 0; i < iterations.length; i++) {
    const row = dataStartRow + i;
    
    // Product Remaining (column 5 = column E)
    const productRemaining = sheet.getRange(row, 5).getValue();
    if (typeof productRemaining === 'number') {
      sheet.getRange(row, 5).setBackground(productRemaining >= 0 ? '#ccffcc' : '#ffcccc');
    }
    
    // Tech Remaining (column 8 = column H)
    const techRemaining = sheet.getRange(row, 8).getValue();
    if (typeof techRemaining === 'number') {
      sheet.getRange(row, 8).setBackground(techRemaining >= 0 ? '#ccffcc' : '#ffcccc');
    }
    
    // Quality Remaining (column 11 = column K)
    const qualityRemaining = sheet.getRange(row, 11).getValue();
    if (typeof qualityRemaining === 'number') {
      sheet.getRange(row, 11).setBackground(qualityRemaining >= 0 ? '#ccffcc' : '#ffcccc');
    }
    
    // Re-apply light grey to Unplanned Work column
    sheet.getRange(row, 12).setBackground('#f5f5f5');
  }
  
  return dataStartRow + iterations.length + 2;
}
// ===== MAIN ORCHESTRATOR FUNCTIONS =====
function createScrumTeamSummary(allIssues, programIncrement, scrumTeam, targetSpreadsheet, sourceSpreadsheet) {
  try {
    console.log(`Creating summary for team: ${scrumTeam}`);
    
    // Filter issues for this team
    const teamIssues = allIssues.filter(issue => 
      (issue.scrumTeam || 'Unassigned') === scrumTeam
    );
    
    if (teamIssues.length === 0) {
      return {
        success: false,
        team: scrumTeam,
        error: `No data found for team ${scrumTeam}`
      };
    }
    
    // Get or create the summary sheet
    const spreadsheet = targetSpreadsheet || SpreadsheetApp.getActiveSpreadsheet();
    // Use source spreadsheet for capacity data if provided
    const capacitySpreadsheet = sourceSpreadsheet || spreadsheet;
    
    const sheetName = `${programIncrement} - ${scrumTeam} Summary`;
    let sheet = spreadsheet.getSheetByName(sheetName);
    
    if (sheet) {
      // Remove existing charts first (sheet.clear() doesn't remove charts)
      const existingCharts = sheet.getCharts();
      existingCharts.forEach(chart => {
        sheet.removeChart(chart);
      });
      console.log(`Removed ${existingCharts.length} existing charts from ${sheetName}`);
      
      // Clear existing content
      sheet.clear();
    } else {
      // Create new sheet
      sheet = spreadsheet.insertSheet(sheetName);
    }
    
    // Clear all existing notes from the sheet more thoroughly
    sheet.clearNotes();
    
    // Set up the sheet
    let currentRow = 1;
    
    // Add title
    sheet.getRange(currentRow, 1).setValue(`${programIncrement} - ${scrumTeam} Summary`);
    sheet.getRange(currentRow, 1).setFontSize(16).setFontWeight('bold').setFontFamily('Comfortaa');
    setRowHeightWithLimit(sheet, currentRow, 30, 70);
    currentRow++;
    
    // Add last refreshed timestamp
    const now = new Date();
    const formattedDate = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    sheet.getRange(currentRow, 1).setValue(`Last Refreshed: ${formattedDate}`);
    sheet.getRange(currentRow, 1).setFontSize(8).setFontStyle('italic').setFontFamily('Comfortaa');
    setRowHeightWithLimit(sheet, currentRow, 20, 70);
    currentRow++;
    
    // Add note about conditional sections
    sheet.getRange(currentRow, 1).setValue('Note: Some sections will only appear when issues are detected (e.g., allocation mismatches, fix version issues)');
    sheet.getRange(currentRow, 1).setFontSize(8).setFontStyle('italic').setFontColor('#666666').setFontFamily('Comfortaa');
    setRowHeightWithLimit(sheet, currentRow, 20, 70);
    currentRow += 2;
    
    // Filter epics and stories for this team
    const epics = teamIssues.filter(i => i.issueType === 'Epic');
    const stories = teamIssues.filter(i => i.issueType !== 'Epic');
    
    // CHANGED ORDER: Planning Progress FIRST
    // Calculate total story points for the gauge
    const totalStoryPoints = calculateTotalStoryPoints(teamIssues, scrumTeam);
    
    // Add Planning Progress section FIRST - pass capacitySpreadsheet for capacity lookups
    currentRow = createTeamPlanningProgressGauges(sheet, currentRow, teamIssues, epics, stories, scrumTeam, programIncrement, totalStoryPoints, capacitySpreadsheet);
    currentRow += 2;
    
    // Add Planned Capacity Distribution chart (from capacity planning sheet)
    currentRow = createPlannedCapacityDistributionChart(sheet, currentRow, scrumTeam, capacitySpreadsheet);
    currentRow += 2;
    
    // Add Allocation Analysis Chart (shows actual planned work vs capacity)
    const allocationResult = createTeamAllocationChart(sheet, currentRow, teamIssues, scrumTeam);
    currentRow = allocationResult.nextRow;
    currentRow += 2;
    
    // Add Epics slotted by Iteration chart
    currentRow = createEpicsSlottedByIteration(sheet, currentRow, teamIssues, scrumTeam, programIncrement);
    currentRow += 2;
    
    // Add Initiative Analysis section (Portfolio & Program Initiative distribution)
    currentRow = createTeamInitiativeAnalysis(sheet, currentRow, teamIssues, scrumTeam, programIncrement);
    currentRow += 2;
    
    // Add All Epics for Planning section
    currentRow = createAllEpicsForPlanning(sheet, currentRow, teamIssues, scrumTeam);
    currentRow += 2;
    
    // Add Release Version Validation section
    currentRow = createReleaseVersionValidation(sheet, currentRow, teamIssues, scrumTeam, programIncrement);
    currentRow += 2;
    
    // Add Allocation Mismatch section (if applicable)
    currentRow = addAllocationMismatchToSummary(sheet, currentRow, teamIssues, scrumTeam);
    
    // Add summary of what was included
    currentRow += 2;
    sheet.getRange(currentRow, 1).setValue('Summary Report Information');
    sheet.getRange(currentRow, 1).setFontSize(10).setFontWeight('bold').setFontFamily('Comfortaa');
    currentRow++;
    
    const endTime = new Date();
    const formattedEndTime = Utilities.formatDate(endTime, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    sheet.getRange(currentRow, 1).setValue(`Report completed at: ${formattedEndTime}`);
    sheet.getRange(currentRow, 1).setFontSize(8).setFontStyle('italic').setFontFamily('Comfortaa');
    currentRow++;
    
    // Set all columns to width 100
    for (let col = 1; col <= 20; col++) {
      sheet.setColumnWidth(col, 100);
    }
    
    // Ensure all rows have reasonable heights (max 70)
    const lastRow = sheet.getLastRow();
    for (let row = 1; row <= lastRow; row++) {
      const currentHeight = sheet.getRowHeight(row);
      if (currentHeight > 70) {
        sheet.setRowHeight(row, 70);
      }
    }
    
    return {
      success: true,
      team: scrumTeam,
      sheetName: sheetName
    };
    
  } catch (error) {
    console.error(`Error creating summary for ${scrumTeam}:`, error);
    return {
      success: false,
      team: scrumTeam,
      error: error.toString()
    };
  }
}

function createScrumTeamSummaries(allIssues, programIncrement, scrumTeams, targetSpreadsheet) {
  const results = {
    success: [],
    failed: [],
    total: scrumTeams.length
  };
  
  scrumTeams.forEach(scrumTeam => {
    try {
      console.log(`Processing team: ${scrumTeam}`);
      
      // Call the singular function for each team
      const result = createScrumTeamSummary(allIssues, programIncrement, scrumTeam, targetSpreadsheet);
      
      if (result.success) {
        results.success.push(result);
        console.log(`Successfully created summary for ${scrumTeam}`);
      } else {
        results.failed.push(result);
        console.log(`Failed to create summary for ${scrumTeam}: ${result.error}`);
      }
    } catch (error) {
      console.error(`Error creating summary for ${scrumTeam}:`, error);
      results.failed.push({
        success: false,
        team: scrumTeam,
        error: error.toString()
      });
    }
  });
  
  console.log(`Summary creation complete: ${results.success.length} succeeded, ${results.failed.length} failed`);
  return results;
}

// ===== SECTION CREATION FUNCTIONS (in order of usage) =====
function calculateTotalStoryPoints(issues, scrumTeam) {
  const stories = issues.filter(i => i.issueType === 'Story' || i.issueType === 'Bug'); // UPDATED: Include Bug type
  let totalStoryPoints = 0;
  
  stories.forEach(story => {
    const category = mapAllocationToCategory(story.allocation);
    if (category) {
      totalStoryPoints += story.storyPoints || 0;
    }
  });
  
  return totalStoryPoints;
}

function createTeamPlanningProgressGauges(sheet, startRow, allIssues, epics, stories, scrumTeam, programIncrement, totalStoryPoints, capacitySpreadsheet) {
  console.log(`Creating planning progress gauges for ${scrumTeam}`);
  
  // Set column widths to 100
  for (let col = 1; col <= 4; col++) {
    sheet.setColumnWidth(col, 100);
  }
  
  // Use capacitySpreadsheet for capacity data if provided, otherwise use sheet's parent
  const spreadsheet = capacitySpreadsheet || sheet.getParent();
  
  // Get capacity data from consolidated capacity sheet
  let percentCapacityUsed = 0;
  let capacityUsedStatus = 'No capacity data';
  
  try {
    if (typeof getCapacityDataForTeamConsolidated === 'function') {
      const capacityData = getCapacityDataForTeamConsolidated(spreadsheet, scrumTeam);
      
      if (capacityData) {
        console.log(`Found capacity data for ${scrumTeam} from consolidated format`);
        
        // Use product capacity (Feature + Compliance) as available capacity
        const availableCapacity = capacityData.productCapacity || 0;
        console.log(`Available capacity (Product): ${availableCapacity}`);
        console.log(`Total story points (calculated): ${totalStoryPoints}`);
        
        if (availableCapacity > 0) {
          // Calculate % of capacity used (total story points / available capacity)
          percentCapacityUsed = Math.round((totalStoryPoints / availableCapacity) * 100);
          
          if (percentCapacityUsed > 100) {
            capacityUsedStatus = `${percentCapacityUsed}% (OVERALLOCATED)`;
          } else {
            capacityUsedStatus = `${percentCapacityUsed}%`;
          }
          
          console.log(`Calculated percentage - Used: ${percentCapacityUsed}%`);
        } else {
          capacityUsedStatus = 'No capacity defined';
          console.log('No available capacity defined for team');
        }
      } else {
        console.log(`No capacity data found for team ${scrumTeam} in consolidated format`);
      }
    } else {
      console.log('getCapacityDataForTeamConsolidated not available');
    }
  } catch (e) {
    console.error(`Error getting capacity data: ${e.message}`);
  }
  
  // Calculate planning metrics for epics
  // UPDATED: Include both Story and Bug types
  const epicStories = stories.filter(s => 
    (s.issueType === 'Story' || s.issueType === 'Bug') &&
    (s.epicLink || s.parentKey)
  );
  const epicsWithAllStoryPoints = new Set();
  
  // Filter out epics that contain "Unplanned" in their summary
  const plannedEpics = epics.filter(epic => 
    !epic.summary || !epic.summary.toLowerCase().includes('unplanned')
  );
  
  plannedEpics.forEach(epic => {
    const epicChildStories = epicStories.filter(s => 
      (s.parentKey === epic.key || s.epicLink === epic.key) &&
      (s.issueType === 'Story' || s.issueType === 'Bug') // UPDATED: Include Bug type
    );
    
    // Check if all stories have story points
    if (epicChildStories.length > 0 && epicChildStories.every(s => s.storyPoints && s.storyPoints > 0)) {
      epicsWithAllStoryPoints.add(epic.key);
    }
  });
  
  const percentEpicsWithStoryPoints = plannedEpics.length > 0 ? 
    Math.round((epicsWithAllStoryPoints.size / plannedEpics.length) * 100) : 0;
  
  // Create the gauge section - no merging, just fill A, B, C with purple
  sheet.getRange(startRow, 1).setValue('Planning Progress');
  sheet.getRange(startRow, 1, 1, 4).setBackground('#E1D5E7');
  sheet.getRange(startRow, 1).setFontSize(14).setFontWeight('bold').setFontColor('black');
  sheet.getRange(startRow, 1).setFontFamily('Comfortaa');
  sheet.getRange(startRow, 1).setVerticalAlignment('middle');
  setRowHeightWithLimit(sheet, startRow, 30, 70);
  startRow++; // No space between title and table
  
  // Planning completion metrics
  const metricsHeaders = ['Metric', '', 'Value', 'Progress'];
  sheet.getRange(startRow, 1, 1, metricsHeaders.length).setValues([metricsHeaders]);
  sheet.getRange(startRow, 1, 1, metricsHeaders.length)
    .setFontWeight('bold')
    .setBackground('#9b7bb8')
    .setFontColor('white')
    .setFontSize(8)
    .setWrap(true)
    .setFontFamily('Comfortaa')
    .setVerticalAlignment('middle');
  
  // Set header row height
  setRowHeightWithLimit(sheet, startRow, 25, 70);
  
  startRow++;
  
  // % of capacity used row
  sheet.getRange(startRow, 1).setValue('% of capacity used');
  sheet.getRange(startRow, 3).setValue(capacityUsedStatus);
  sheet.getRange(startRow, 1, 1, 4).setFontSize(8).setWrap(true).setFontFamily('Comfortaa').setVerticalAlignment('middle');
  
  // Set row height
  setRowHeightWithLimit(sheet, startRow, 25, 70);
  
  // Create progress bar for capacity used
  createProgressBar(sheet, startRow, 4, percentCapacityUsed);
  
  // If overallocated, make the value cell red too
  if (percentCapacityUsed > 100) {
    sheet.getRange(startRow, 3).setFontColor('#ff0000').setFontWeight('bold');
  }
  
  startRow++;
  
  // Epics with story points row
  sheet.getRange(startRow, 1).setValue('% of Epics with All Stories Pointed');
  sheet.getRange(startRow, 3).setValue(`${percentEpicsWithStoryPoints}%`);
  sheet.getRange(startRow, 1, 1, 4).setFontSize(8).setWrap(true).setFontFamily('Comfortaa').setVerticalAlignment('middle');
  
  // Set row height
  setRowHeightWithLimit(sheet, startRow, 25, 70);
  
  // Create progress bar for epics
  createProgressBar(sheet, startRow, 4, percentEpicsWithStoryPoints);
  startRow += 2;
  
  // Add Iteration Slotting chart below the metrics - use capacitySpreadsheet for capacity data
  const iterationSlottingEndRow = createIterationSlottingChart(
    sheet, 
    startRow,
    allIssues, 
    scrumTeam, 
    programIncrement,
    spreadsheet  // This is capacitySpreadsheet (source) for capacity lookups
  );
  
  // Add Role Breakdown chart below Iteration Slotting (if roles exist for this team)
  const roleBreakdownEndRow = createRoleBreakdownChart(
    sheet,
    iterationSlottingEndRow,
    allIssues,
    scrumTeam,
    programIncrement,
    spreadsheet  // This is capacitySpreadsheet (source) for capacity lookups
  );
  
  return roleBreakdownEndRow;
}

/**
 * Creates a Planned Capacity Distribution chart showing how capacity is allocated
 * across allocation types (Product, Tech/Platform, Quality, KLO) from the capacity planning sheet
 * 
 * @param {Sheet} sheet - The target sheet
 * @param {number} startRow - Starting row
 * @param {string} scrumTeam - Team name
 * @param {Spreadsheet} capacitySpreadsheet - Spreadsheet containing capacity data
 * @returns {number} Next available row after the chart
 */
function createPlannedCapacityDistributionChart(sheet, startRow, scrumTeam, capacitySpreadsheet) {
  console.log(`Creating Planned Capacity Distribution chart for ${scrumTeam}`);
  
  // Get capacity data from consolidated capacity sheet
  let capacityData = null;
  try {
    if (typeof getCapacityDataForTeamConsolidated === 'function') {
      capacityData = getCapacityDataForTeamConsolidated(capacitySpreadsheet, scrumTeam);
    }
  } catch (e) {
    console.log(`Error getting capacity data: ${e.message}`);
  }
  
  if (!capacityData || !capacityData.allocations) {
    console.log(`No capacity data found for ${scrumTeam} - skipping distribution chart`);
    return startRow;
  }
  
  // Title
  sheet.getRange(startRow, 1).setValue('Planned Capacity Distribution');
  sheet.getRange(startRow, 1, 1, 6).setBackground('#E1D5E7');
  sheet.getRange(startRow, 1).setFontSize(14).setFontWeight('bold').setFontColor('black');
  sheet.getRange(startRow, 1).setFontFamily('Comfortaa');
  sheet.getRange(startRow, 1).setVerticalAlignment('middle');
  setRowHeightWithLimit(sheet, startRow, 30, 70);
  startRow++;
  
  // Headers
  const headers = ['Allocation Type', 'Planned Capacity', '% of Total', '', '', ''];
  sheet.getRange(startRow, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(startRow, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#9b7bb8')
    .setFontColor('white')
    .setFontSize(10)
    .setFontFamily('Comfortaa');
  startRow++;
  
  // Calculate totals
  const allocs = capacityData.allocations;
  const productCapacity = (allocs.productFeature || 0) + (allocs.productCompliance || 0);
  const techCapacity = allocs.techPlatform || 0;
  const qualityCapacity = allocs.quality || 0;
  const kloCapacity = allocs.klo || 0;
  const unplannedCapacity = allocs.unplannedWork || 0;
  
  const totalCapacity = productCapacity + techCapacity + qualityCapacity + kloCapacity + unplannedCapacity;
  
  // Helper to calculate percentage
  const calcPercent = (value) => {
    if (totalCapacity === 0) return '0%';
    return Math.round((value / totalCapacity) * 100) + '%';
  };
  
  // Data rows with colors
  const dataRows = [
    { name: 'Product (Feature + Compliance)', value: productCapacity, color: '#1B365D' },
    { name: 'Tech / Platform', value: techCapacity, color: '#6B3FA0' },
    { name: 'Quality', value: qualityCapacity, color: '#4285F4' },
    { name: 'KLO (Keep Lights On)', value: kloCapacity, color: '#FFC72C' },
    { name: 'Unplanned Work', value: unplannedCapacity, color: '#9AA0A6' }
  ];
  
  const dataStartRow = startRow;
  
  dataRows.forEach((row, index) => {
    sheet.getRange(startRow, 1).setValue(row.name);
    sheet.getRange(startRow, 2).setValue(Math.round(row.value));
    sheet.getRange(startRow, 3).setValue(calcPercent(row.value));
    
    // Color indicator in column 4
    sheet.getRange(startRow, 4).setBackground(row.color);
    sheet.getRange(startRow, 4).setValue('');
    
    // Alternate row colors
    if (index % 2 === 1) {
      sheet.getRange(startRow, 1, 1, 3).setBackground('#f5f5f5');
    }
    
    startRow++;
  });
  
  // Total row
  sheet.getRange(startRow, 1).setValue('TOTAL');
  sheet.getRange(startRow, 2).setValue(Math.round(totalCapacity));
  sheet.getRange(startRow, 3).setValue('100%');
  sheet.getRange(startRow, 1, 1, 3)
    .setFontWeight('bold')
    .setBackground('#e0e0e0');
  startRow++;
  
  // Create pie chart
  try {
    const labelRange = sheet.getRange(dataStartRow, 1, dataRows.length, 1);
    const valueRange = sheet.getRange(dataStartRow, 2, dataRows.length, 1);
    
    const chart = sheet.newChart()
      .setChartType(Charts.ChartType.PIE)
      .addRange(labelRange)
      .addRange(valueRange)
      .setPosition(dataStartRow, 6, 0, 0)
      .setOption('title', 'Capacity Distribution')
      .setOption('width', 350)
      .setOption('height', 200)
      .setOption('pieSliceText', 'percentage')
      .setOption('legend', { position: 'right', textStyle: { fontSize: 9 } })
      .setOption('titleTextStyle', { fontSize: 11, bold: true })
      .setOption('colors', dataRows.map(r => r.color))
      .build();
    
    sheet.insertChart(chart);
  } catch (chartError) {
    console.log(`Could not create pie chart: ${chartError.message}`);
  }
  
  return startRow + 1;
}

function createTeamAllocationChart(sheet, startRow, issues, scrumTeam) {
  console.log(`Creating allocation chart for team: ${scrumTeam}`);
  
  const epics = issues.filter(i => i.issueType === 'Epic');
  const stories = issues.filter(i => i.issueType !== 'Epic');
  
  // Title for chart section - Clear any existing formatting and fill ALL columns (A through O) with purple
  sheet.getRange(startRow, 1, 1, 15).clearFormat(); // Clear any existing formatting first
  sheet.getRange(startRow, 1).setValue('Allocation Analysis Chart');
  sheet.getRange(startRow, 1, 1, 15).setBackground('#E1D5E7'); // Fill all 15 columns (A through O) with purple
  sheet.getRange(startRow, 1).setFontSize(14).setFontWeight('bold').setFontColor('black');
  sheet.getRange(startRow, 1).setFontFamily('Comfortaa');
  sheet.getRange(startRow, 1).setVerticalAlignment('middle');
  setRowHeightWithLimit(sheet, startRow, 30, 70);
  startRow++; // No space between title and table
  
  // Updated headers with new column names and additional columns
  const chartHeaders = [
    'Allocation Type', '', 'Baseline Capacity', 'PLANNED Allocation',
    'PLANNED Availability', '% PLANNED Availability', 'Story points used for PI',
    'Used Availability', '% still available for PI',
    'Capacity prior to Code Freeze', 'Remaining for Use overall prior to Code Freeze', '% availability prior to code freeze',
    'Slotted code freeze capacity', 'Slotted Code Freeze Availability', '% Code Freeze Availability'
  ];
  
  // Write headers
  sheet.getRange(startRow, 1, 1, chartHeaders.length).setValues([chartHeaders]);
  
  // Format header row - CHANGED TO PURPLE (#9b7bb8)
  sheet.getRange(startRow, 1, 1, chartHeaders.length)
    .setFontWeight('bold')
    .setBackground('#9b7bb8')  // Purple color
    .setFontColor('white')
    .setFontSize(8)
    .setWrap(true)
    .setFontFamily('Comfortaa')
    .setVerticalAlignment('middle');
  
  // Set header row height for wrapped text (max 70)
  setRowHeightWithLimit(sheet, startRow, 50, 70);
  
  // Center align headers for columns C through O (columns 3-15)
  sheet.getRange(startRow, 3, 1, chartHeaders.length - 2).setHorizontalAlignment('center');
  
  // Add notes to headers (these will be re-added each refresh after clearing)
  sheet.getRange(startRow, 3).setNote('Comes from Capacity planning sheet');
  sheet.getRange(startRow, 13).setNote('Looking at sprint assigned');
  sheet.getRange(startRow, 15).setNote('Looking at sprint assigned');
  
  startRow++;
  
  // Extract PI number from the sheet name
  const sheetName = sheet.getName();
  const piMatch = sheetName.match(/PI (\d+)/);
  const piNumber = piMatch ? parseInt(piMatch[1]) : null;
  const isPI11or12 = piNumber === 11 || piNumber === 12;
  
  // Define allocation categories
  const allocations = [
    { name: 'Features (Product - Compliance & Feature)', capacityColumn: 'B', columnIndex: 2 },
    { name: 'Tech / Platform', capacityColumn: 'C', columnIndex: 3 },
    { name: 'Planned KLO', capacityColumn: 'D', columnIndex: 4 },
    { name: 'Planned Quality', capacityColumn: 'E', columnIndex: 5 },
    { name: 'Unplanned Quality', capacityColumn: 'F', columnIndex: 6, isUnplanned: true }
  ];
  
  // Calculate anticipated and current allocations
  const allocationData = {};
  allocations.forEach(alloc => {
    allocationData[alloc.name] = {
      plannedAllocation: 0,
      usedCapacity: 0,
      slottedCodeFreezeCapacity: 0
    };
  });
  
  // Sum up allocations
  epics.forEach(epic => {
    const category = mapAllocationToCategory(epic.allocation); // mapAllocationToCategory is defined in the main config file
    if (allocationData[category]) {
      const featurePointValue = (epic.featurePoints || 0) * 10;
      allocationData[category].plannedAllocation += featurePointValue;
    }
  });
  
  stories.forEach(story => {
    const category = mapAllocationToCategory(story.allocation); // mapAllocationToCategory is defined in the main config file
    if (allocationData[category]) {
      const storyPointValue = story.storyPoints || 0;
      allocationData[category].usedCapacity += storyPointValue;
      
      // Check if sprint meets code freeze criteria
      if (story.sprintName) {
        // Try multiple patterns to match sprint names
        const patterns = [
          new RegExp(`PI\\s*${piNumber}\\s*\\.\\s*([1-5])`, 'i'),  // "PI 12.1" format
          new RegExp(`${piNumber}\\s*\\.\\s*([1-5])`, 'i'),        // "12.1" format
          new RegExp(`\\b${piNumber}\\.([1-5])\\b`)              // Strict "12.1" format
        ];
        
        let matched = false;
        for (const pattern of patterns) {
          const match = story.sprintName.match(pattern);
          if (match) {
            const iteration = parseInt(match[1]);
            if (iteration >= 1 && iteration <= 5) {
              allocationData[category].slottedCodeFreezeCapacity += storyPointValue;
              matched = true;
              break;
            }
          }
        }
        
        if (!matched && story.sprintName.toLowerCase().includes(`${piNumber}.`)) {
          console.log(`Sprint not matched for slotted capacity: ${story.sprintName}`);
        }
      }
    }
  });
  
  // Add borders for the sections
  const borderStartRow = startRow;
  const borderEndRow = startRow + allocations.length;
  
  // Feature point allocation frame
  sheet.getRange(borderStartRow - 1, 4, allocations.length + 1, 3).setBorder(
    true, true, true, true, false, false, '#555555', SpreadsheetApp.BorderStyle.SOLID_THICK
  );
  
  // Ticket/Story point break down frame
  sheet.getRange(borderStartRow - 1, 7, allocations.length + 1, 3).setBorder(
    true, true, true, true, false, false, '#555555', SpreadsheetApp.BorderStyle.SOLID_THICK
  );
  
  // All iteration allocation frame (was "All Story points")
  sheet.getRange(borderStartRow - 1, 10, allocations.length + 1, 3).setBorder(
    true, true, true, true, false, false, '#555555', SpreadsheetApp.BorderStyle.SOLID_THICK
  );
  
  // Iteration allocation prior to code freeze frame (was "Sprint prior to Code Freeze")
  sheet.getRange(borderStartRow - 1, 13, allocations.length + 1, 3).setBorder(
    true, true, true, true, false, false, '#555555', SpreadsheetApp.BorderStyle.SOLID_THICK
  );
  
  // Write data rows
  allocations.forEach((alloc, index) => {
    const rowNum = startRow + index;
    const data = allocationData[alloc.name];
    
    // Apply alternating row background - white and light grey for entire row
    if (index % 2 === 0) {
      sheet.getRange(rowNum, 1, 1, chartHeaders.length).setBackground('#ffffff'); // White for even rows
    } else {
      sheet.getRange(rowNum, 1, 1, chartHeaders.length).setBackground('#f5f5f5'); // Light grey for odd rows
    }
    
    // Set allocation type - NO WRAP for column A
    sheet.getRange(rowNum, 1).setValue(alloc.name);
    sheet.getRange(rowNum, 1).setWrap(false);  // Turn off text wrapping
    
    // Set FORMULA for planned capacity with flexible team name matching
    // This formula handles case differences and attempts to match team names flexibly
    const capacityFormula = `=IFERROR(INDEX(Capacity!${alloc.capacityColumn}:${alloc.capacityColumn},MATCH(UPPER(TRIM("${scrumTeam}")),ARRAYFORMULA(UPPER(TRIM(Capacity!A:A))),0)),0)`;
    sheet.getRange(rowNum, 3).setFormula(capacityFormula);
    
    // Handle Unplanned Quality row differently
    if (alloc.isUnplanned) {
      // Leave all columns blank from D to O but maintain row background
      sheet.getRange(rowNum, 4, 1, 12).setValue('');
    } else {
      // Normal row processing
      sheet.getRange(rowNum, 4).setValue(data.plannedAllocation);
      
      sheet.getRange(rowNum, 5).setFormula(`=C${rowNum}-D${rowNum}`);
      sheet.getRange(rowNum, 6).setFormula(`=IF(C${rowNum}>0,ROUND(E${rowNum}/C${rowNum}*100,0)&"%","0%")`);
      
      sheet.getRange(rowNum, 7).setValue(data.usedCapacity);
      sheet.getRange(rowNum, 8).setFormula(`=C${rowNum}-G${rowNum}`);
      sheet.getRange(rowNum, 9).setFormula(`=IF(C${rowNum}>0,ROUND(H${rowNum}/C${rowNum}*100,0)&"%","0%")`);
      
      // Column J formulas based on allocation type
      if (alloc.name === 'Features (Product - Compliance & Feature)') {
        sheet.getRange(rowNum, 10).setFormula(`=C18`);
      } else if (alloc.name === 'Tech / Platform') {
        sheet.getRange(rowNum, 10).setFormula(`=F18`);
      } else if (alloc.name === 'Planned KLO') {
        sheet.getRange(rowNum, 10).setFormula(`=ROUND((C28/13)*8,0)`);
      } else if (alloc.name === 'Planned Quality') {
        sheet.getRange(rowNum, 10).setFormula(`=I18`);
      } else {
        // For other rows, use the old formula
        const codeFreezeMultiplier = isPI11or12 ? 8 : 10;
        sheet.getRange(rowNum, 10).setFormula(`=ROUND(C${rowNum}/13*${codeFreezeMultiplier},0)`);
      }
      
      sheet.getRange(rowNum, 11).setFormula(`=J${rowNum}-G${rowNum}`);
      sheet.getRange(rowNum, 12).setFormula(`=IF(J${rowNum}>0,ROUND(K${rowNum}/J${rowNum}*100,0)&"%","0%")`);
      
      // New columns for slotted code freeze
      sheet.getRange(rowNum, 13).setValue(data.slottedCodeFreezeCapacity);
      sheet.getRange(rowNum, 14).setFormula(`=J${rowNum}-M${rowNum}`);
      sheet.getRange(rowNum, 15).setFormula(`=IF(J${rowNum}>0,ROUND(N${rowNum}/J${rowNum}*100,0)&"%","0%")`);
      
      // Override columns D and F with darker grey (always)
      sheet.getRange(rowNum, 4).setBackground('#d3d3d3'); // Column D - dark grey
      sheet.getRange(rowNum, 6).setBackground('#d3d3d3'); // Column F - dark grey
    }
  });
  
  // Set font size for data rows and wrap text (except column A)
  sheet.getRange(startRow, 1, allocations.length, 1).setFontSize(8).setWrap(false).setFontFamily('Comfortaa').setVerticalAlignment('middle');
  sheet.getRange(startRow, 2, allocations.length, chartHeaders.length - 1).setFontSize(8).setWrap(true).setFontFamily('Comfortaa').setVerticalAlignment('middle');
  
  // IMPORTANT: Set number format for columns D and E (and other numeric columns) to prevent % auto-formatting
  sheet.getRange(startRow, 3, allocations.length, 1).setNumberFormat('0');  // Column C - Baseline Capacity
  sheet.getRange(startRow, 4, allocations.length, 1).setNumberFormat('0');  // Column D - PLANNED Allocation
  sheet.getRange(startRow, 5, allocations.length, 1).setNumberFormat('0');  // Column E - PLANNED Availability
  sheet.getRange(startRow, 7, allocations.length, 1).setNumberFormat('0');  // Column G - Story points planned
  sheet.getRange(startRow, 8, allocations.length, 1).setNumberFormat('0');  // Column H - Current Availability
  sheet.getRange(startRow, 10, allocations.length, 1).setNumberFormat('0'); // Column J - Capacity prior to CF
  sheet.getRange(startRow, 11, allocations.length, 1).setNumberFormat('0'); // Column K - Remaining for Use
  sheet.getRange(startRow, 13, allocations.length, 1).setNumberFormat('0'); // Column M - Slotted CF capacity
  sheet.getRange(startRow, 14, allocations.length, 1).setNumberFormat('0'); // Column N - Slotted CF Availability
  
  // Set standard row height for data rows
  for (let i = 0; i < allocations.length; i++) {
    setRowHeightWithLimit(sheet, startRow + i, 25, 70);
  }
  
  // Center align all data columns
  sheet.getRange(startRow, 3, allocations.length, 13).setHorizontalAlignment('center');
  
  // Apply light blue background to column C for all data rows
  sheet.getRange(startRow, 3, allocations.length, 1).setBackground('#e6f2ff');
  
  // Apply light blue background to column J for data rows (J26:J30)
  sheet.getRange(startRow, 10, allocations.length, 1).setBackground('#e6f2ff');
  
  // Apply conditional formatting to variance cells AFTER row backgrounds
  for (let i = 0; i < allocations.length; i++) {
    const row = startRow + i;
    const alloc = allocations[i];
    
    if (!alloc.isUnplanned) {
      SpreadsheetApp.flush();
      
      const plannedValue = sheet.getRange(row, 5).getValue();
      if (typeof plannedValue === 'number') {
        sheet.getRange(row, 5).setBackground(plannedValue >= 0 ? '#ccffcc' : '#ffcccc');
      }
      
      const usedValue = sheet.getRange(row, 8).getValue();
      if (typeof usedValue === 'number') {
        sheet.getRange(row, 8).setBackground(usedValue >= 0 ? '#ccffcc' : '#ffcccc');
      }
      
      const remainingValue = sheet.getRange(row, 11).getValue();
      if (typeof remainingValue === 'number') {
        sheet.getRange(row, 11).setBackground(remainingValue >= 0 ? '#ccffcc' : '#ffcccc');
      }
      
      const slottedValue = sheet.getRange(row, 14).getValue();
      if (typeof slottedValue === 'number') {
        sheet.getRange(row, 14).setBackground(slottedValue >= 0 ? '#ccffcc' : '#ffcccc');
      }
    }
  }
  
  // Set all column widths to 100
  for (let col = 1; col <= chartHeaders.length; col++) {
    sheet.setColumnWidth(col, 100);
  }
  
  startRow += allocations.length;
  
  // Add TOTALS row - using darker grey background
  const totalRow = startRow + 1;  // Skip one row as requested
  
  // Set TOTAL in column A
  sheet.getRange(totalRow, 1).setValue('TOTAL');
  sheet.getRange(totalRow, 1).setWrap(false);  // No wrap for TOTAL cell
  
  // Apply darker grey background to columns D through O in totals row
  sheet.getRange(totalRow, 4, 1, 12).setBackground('#808080'); // Dark gray for D through O
  
  // Keep column B and C with lighter grey
  sheet.getRange(totalRow, 2, 1, 2).setBackground('#d3d3d3');
  
  // Fixed formulas for totals
  const firstDataRow = totalRow - allocations.length - 1;
  const lastDataRow = totalRow - 2;
  
  // Sum formulas for numeric columns only (not percentage columns)
  sheet.getRange(totalRow, 3).setFormula(`=SUM(C${firstDataRow}:C${lastDataRow})`); // Planned Capacity
  sheet.getRange(totalRow, 4).setFormula(`=SUM(D${firstDataRow}:D${lastDataRow - 1})`); // PLANNED Allocation (exclude unplanned)    
  sheet.getRange(totalRow, 7).setFormula(`=SUM(G${firstDataRow}:G${lastDataRow})`); // Story points used for PI
  sheet.getRange(totalRow, 10).setFormula(`=SUM(J${firstDataRow}:J${lastDataRow - 1})`); // Capacity prior to Code Freeze (exclude unplanned)
  sheet.getRange(totalRow, 13).setFormula(`=SUM(M${firstDataRow}:M${lastDataRow})`); // Slotted code freeze capacity
  
  // Leave percentage and variance columns blank
  sheet.getRange(totalRow, 5).setValue(''); // PLANNED Availability
  sheet.getRange(totalRow, 6).setValue(''); // % PLANNED Availability
  sheet.getRange(totalRow, 8).setValue(''); // Used Availability
  sheet.getRange(totalRow, 9).setValue(''); // % still available
  sheet.getRange(totalRow, 11).setValue(''); // Remaining overall
  sheet.getRange(totalRow, 12).setValue(''); // % availability
  sheet.getRange(totalRow, 14).setValue(''); // Slotted Code Freeze Availability
  sheet.getRange(totalRow, 15).setValue(''); // % Code Freeze Availability  
  // Format TOTALS row
  sheet.getRange(totalRow, 1, 1, chartHeaders.length).setFontWeight('bold').setFontSize(8).setFontFamily('Comfortaa').setVerticalAlignment('middle');
  sheet.getRange(totalRow, 3, 1, 13).setHorizontalAlignment('center');
  sheet.getRange(totalRow, 3).setNumberFormat('0'); // Ensure numeric format
  sheet.getRange(totalRow, 4).setNumberFormat('0');
  sheet.getRange(totalRow, 7).setNumberFormat('0');
  sheet.getRange(totalRow, 10).setNumberFormat('0');
  sheet.getRange(totalRow, 13).setNumberFormat('0');
  
  // Apply light blue background to column C for totals row
  sheet.getRange(totalRow, 3).setBackground('#e6f2ff');
  
  // Apply light blue background to column J for totals row
  sheet.getRange(totalRow, 10).setBackground('#e6f2ff');
  
  // Set row height for totals row
  setRowHeightWithLimit(sheet, totalRow, 25, 70);
  
  // CREATE NAMED RANGE for total story points (G column of totals row)
  const spreadsheet = sheet.getParent();
  const storyPointsCell = sheet.getRange(totalRow, 7); // Column G of totals row
  
  // Create a unique named range for this team
  const namedRangeName = `TotalStoryPoints_${scrumTeam.replace(/[\s-]/g, '_')}`;
  
  // Remove existing named range if it exists
  try {
    const existingRange = spreadsheet.getRangeByName(namedRangeName);
    if (existingRange) {
      spreadsheet.removeNamedRange(namedRangeName);
    }
  } catch (e) {
    // Named range doesn't exist, that's fine
  }
  
  // Create the new named range
  spreadsheet.setNamedRange(namedRangeName, storyPointsCell);
  console.log(`Created named range ${namedRangeName} for cell G${totalRow}`);
  
  // Return both the next row and the totals row number
  return {
    nextRow: totalRow + 1,
    totalsRow: totalRow
  };
}


function createEpicsSlottedByIteration(sheet, startRow, issues, scrumTeam, programIncrement) {
  console.log(`Creating Epics slotted by Iteration chart for ${scrumTeam}`);
  
  // Extract PI number
  const piNumber = parseInt(programIncrement.replace('PI ', ''));
  if (isNaN(piNumber)) {
    console.error('Invalid PI number in programIncrement:', programIncrement);
    return startRow;
  }
  
  // Filter epics and stories/bugs - UPDATED
  const epics = issues.filter(i => i.issueType === 'Epic');
  const stories = issues.filter(i => i.issueType === 'Story' || i.issueType === 'Bug'); // UPDATED: Include Bug type
  
  if (epics.length === 0) {
    console.log('No epics found for Epics slotted by Iteration chart');
    return startRow;
  }
  
  // Title - fill columns A through K with purple
  sheet.getRange(startRow, 1).setValue('Epics slotted by Iteration');
  sheet.getRange(startRow, 1, 1, 11).setBackground('#E1D5E7');
  sheet.getRange(startRow, 1).setFontSize(14).setFontWeight('bold').setFontColor('black');
  sheet.getRange(startRow, 1).setFontFamily('Comfortaa');
  sheet.getRange(startRow, 1).setVerticalAlignment('middle');
  setRowHeightWithLimit(sheet, startRow, 30, 70);
  
  const headerRow = startRow + 1;
  
  // Headers
  const headers = [
    'Key', 'Summary', 'Iteration 1', 'Iteration 2', 'Iteration 3', 
    'Iteration 4', 'Iteration 5', 'Iteration 6', 'Total Slotted', 
    'Unslotted Points', 'Tickets Without Points'
  ];
  
  sheet.getRange(headerRow, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(headerRow, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#9b7bb8')
    .setFontColor('white')
    .setFontSize(8)
    .setWrap(true)
    .setFontFamily('Comfortaa')
    .setVerticalAlignment('middle');
  
  setRowHeightWithLimit(sheet, headerRow, 40, 70);
  
  // Set column widths
  sheet.setColumnWidth(1, 100);  // Key
  sheet.setColumnWidth(2, 300);  // Summary (wider)
  for (let col = 3; col <= headers.length; col++) {
    sheet.setColumnWidth(col, 100);
  }
  
  const dataStartRow = headerRow + 1;
  
  // Process each epic
  const epicData = [];
  
  epics.forEach(epic => {
    // Find stories for this epic
    const epicStories = stories.filter(s => 
      s.parentKey === epic.key || s.epicLink === epic.key
    );
    
    // Initialize iteration data
    const iterationData = {
      1: 0, 2: 0, 3: 0, 4: 0, 5: 0, 6: 0,
      unslotted: 0,
      ticketsWithoutPoints: 0,
      totalSlotted: 0
    };
    
    // Process each story
    epicStories.forEach(story => {
      // Count stories without points
      if (!story.storyPoints || story.storyPoints === 0) {
        iterationData.ticketsWithoutPoints++;
        return;
      }
      
      const storyPoints = story.storyPoints;
      let slotted = false;
      
      // Check if story has a sprint
      if (story.sprintName) {
        // Try multiple patterns to match sprint names (same as calculateSlottedData)
        const patterns = [
          new RegExp(`PI\\s*${piNumber}\\s*\\.\\s*(\\d)`, 'i'),  // "PI 12.1" format
          new RegExp(`${piNumber}\\s*\\.\\s*(\\d)`, 'i'),        // "12.1" format
          new RegExp(`\\b${piNumber}\\.(\\d)\\b`)                // Strict "12.1" format
        ];
        
        for (const pattern of patterns) {
          const match = story.sprintName.match(pattern);
          if (match) {
            const iteration = parseInt(match[1]);
            if (iteration >= 1 && iteration <= 6) {
              iterationData[iteration] += storyPoints;
              iterationData.totalSlotted += storyPoints;
              slotted = true;
              break;
            }
          }
        }
      }
      
      // If not slotted to an iteration, count as unslotted
      if (!slotted) {
        iterationData.unslotted += storyPoints;
      }
    });
    
    epicData.push({
      key: epic.key,
      summary: epic.summary || '',
      iterations: iterationData,
      url: epic.url,
      costOfDelay: parseCostOfDelay(epic.costOfDelay) // For sorting
    });
  });
  
  // Sort by Cost of Delay (descending) to match All Epics for Planning order
  epicData.sort((a, b) => b.costOfDelay - a.costOfDelay);
  
  // Write data rows
  epicData.forEach((epic, index) => {
    const rowNum = dataStartRow + index;
    
    // Apply alternating row colors
    if (index % 2 === 0) {
      sheet.getRange(rowNum, 1, 1, headers.length).setBackground('#ffffff');
    } else {
      sheet.getRange(rowNum, 1, 1, headers.length).setBackground('#f5f5f5');
    }
    
    // Key (with hyperlink) - JIRA_CONFIG.baseUrl is defined in the main config file
    const epicKeyFormula = `=HYPERLINK("${JIRA_CONFIG.baseUrl}/browse/${epic.key}","${epic.key}")`;
    sheet.getRange(rowNum, 1).setFormula(epicKeyFormula);
    
    // Summary
    sheet.getRange(rowNum, 2).setValue(epic.summary);
    sheet.getRange(rowNum, 2).setWrap(true);
    
    // Iteration columns (3-8)
    for (let iter = 1; iter <= 6; iter++) {
      const value = epic.iterations[iter];
      if (value > 0) {
        sheet.getRange(rowNum, 2 + iter).setValue(value);
      }
      // Apply light green background if has points
      if (value > 0) {
        sheet.getRange(rowNum, 2 + iter).setBackground('#e6ffe6');
      }
    }
    
    // Total Slotted (column 9)
    sheet.getRange(rowNum, 9).setValue(epic.iterations.totalSlotted);
    sheet.getRange(rowNum, 9).setFontWeight('bold');
    
    // Unslotted Points (column 10)
    if (epic.iterations.unslotted > 0) {
      sheet.getRange(rowNum, 10).setValue(epic.iterations.unslotted);
      sheet.getRange(rowNum, 10).setBackground('#fff3cd'); // Light yellow warning
    }
    
    // Tickets Without Points (column 11)
    if (epic.iterations.ticketsWithoutPoints > 0) {
      sheet.getRange(rowNum, 11).setValue(epic.iterations.ticketsWithoutPoints);
      sheet.getRange(rowNum, 11).setBackground('#ffcccc'); // Light red warning
    }
  });
  
  // Format data area
  if (epicData.length > 0) {
    sheet.getRange(dataStartRow, 1, epicData.length, headers.length)
      .setFontSize(8)
      .setWrap(true)
      .setFontFamily('Comfortaa')
      .setVerticalAlignment('middle');
    
    // Center align numeric columns (3-11)
    sheet.getRange(dataStartRow, 3, epicData.length, 9).setHorizontalAlignment('center');
    
    // Set reasonable row heights
    for (let i = 0; i < epicData.length; i++) {
      setRowHeightWithLimit(sheet, dataStartRow + i, 30, 70);
    }
  }
  
  // Add totals row
  const totalsRow = dataStartRow + epicData.length + 1;
  
  // Calculate totals
  const totals = {
    1: 0, 2: 0, 3: 0, 4: 0, 5: 0, 6: 0,
    totalSlotted: 0,
    unslotted: 0,
    ticketsWithoutPoints: 0
  };
  
  epicData.forEach(epic => {
    for (let iter = 1; iter <= 6; iter++) {
      totals[iter] += epic.iterations[iter];
    }
    totals.totalSlotted += epic.iterations.totalSlotted;
    totals.unslotted += epic.iterations.unslotted;
    totals.ticketsWithoutPoints += epic.iterations.ticketsWithoutPoints;
  });
  
  // Write totals row
  sheet.getRange(totalsRow, 1).setValue('TOTALS');
  sheet.getRange(totalsRow, 1, 1, 2).merge();
  sheet.getRange(totalsRow, 1).setHorizontalAlignment('right');
  
  // Write total values
  for (let iter = 1; iter <= 6; iter++) {
    sheet.getRange(totalsRow, 2 + iter).setValue(totals[iter]);
  }
  sheet.getRange(totalsRow, 9).setValue(totals.totalSlotted);
  sheet.getRange(totalsRow, 10).setValue(totals.unslotted);
  sheet.getRange(totalsRow, 11).setValue(totals.ticketsWithoutPoints);
  
  // Format totals row
  sheet.getRange(totalsRow, 1, 1, headers.length)
    .setFontWeight('bold')
    .setFontSize(8)
    .setBackground('#e0e0e0')
    .setFontFamily('Comfortaa')
    .setVerticalAlignment('middle');
  
  sheet.getRange(totalsRow, 3, 1, 9).setHorizontalAlignment('center');
  
  setRowHeightWithLimit(sheet, totalsRow, 25, 70);
  
  // Add borders around the data section
  sheet.getRange(headerRow, 1, epicData.length + 2, headers.length).setBorder(
    true, true, true, true, false, false,
    '#666666', SpreadsheetApp.BorderStyle.SOLID_MEDIUM
  );
  
  // Add thick borders around iteration columns
  sheet.getRange(headerRow, 3, epicData.length + 2, 6).setBorder(
    true, true, true, true, false, false,
    '#555555', SpreadsheetApp.BorderStyle.SOLID_THICK
  );
  
  return totalsRow + 2;
}

function createAllEpicsForPlanning(sheet, startRow, issues, scrumTeam) {
  console.log(`Creating All Epics for Planning section for ${scrumTeam}`);
  
  // Filter epics and stories/bugs - UPDATED
  const epics = issues.filter(i => i.issueType === 'Epic');
  const stories = issues.filter(i => i.issueType === 'Story' || i.issueType === 'Bug'); // UPDATED: Include Bug type
  
  if (epics.length === 0) {
    console.log('No epics found for All Epics section');
    return startRow;
  }
  
  // Section title - no merge, fill A, B, C with purple
  sheet.getRange(startRow, 1).setValue('All Epics for Planning');
  sheet.getRange(startRow, 1, 1, 14).setBackground('#E1D5E7');
  sheet.getRange(startRow, 1).setFontSize(14).setFontWeight('bold').setFontColor('black');
  sheet.getRange(startRow, 1).setFontFamily('Comfortaa');
  sheet.getRange(startRow, 1).setVerticalAlignment('middle');
  setRowHeightWithLimit(sheet, startRow, 30, 70);
  startRow++; // No space between title and table
  
  // Updated headers - all columns visible
  const headers = [
    'Key', 'Summary', 'Cost of Delay', 'Feature Points', 'Feature Point Conversion Value',
    'Ticket Count', 'Story Points', 'Sized %', 'Allocation', 'Sprint',
    'Fix Version', 'Status', 'Iteration Start (PI #)', 'Iteration End (PI #)'
  ];
  
  sheet.getRange(startRow, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(startRow, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#9b7bb8')
    .setFontColor('white')
    .setFontSize(8)
    .setWrap(true)
    .setFontFamily('Comfortaa')
    .setVerticalAlignment('middle');
  
  // Add note to Sized % column
  sheet.getRange(startRow, 8).setNote('Percentage of tickets with story points assigned');
  
  // Set header row height
  setRowHeightWithLimit(sheet, startRow, 40, 70);
  
  // Set column widths - Summary column wider
  sheet.setColumnWidth(1, 100);  // Key
  sheet.setColumnWidth(2, 300);  // Summary (wider)
  for (let col = 3; col <= headers.length; col++) {
    sheet.setColumnWidth(col, 100);
  }
  
  startRow++;
  
  // Process each epic
  const epicData = [];
  epics.forEach(epic => {
    // Find stories and bugs for this epic - UPDATED
    const epicStories = stories.filter(s => 
      s.parentKey === epic.key || s.epicLink === epic.key
    );
    
    const storyCount = epicStories.length;
    const totalStoryPoints = epicStories.reduce((sum, s) => sum + (s.storyPoints || 0), 0);
    const storiesWithPoints = epicStories.filter(s => s.storyPoints && s.storyPoints > 0).length;
    const sizedPercent = storyCount > 0 ? Math.round((storiesWithPoints / storyCount) * 100) : 0;
    
    epicData.push({
      key: epic.key,
      summary: epic.summary || '',
      costOfDelay: parseCostOfDelay(epic.costOfDelay),
      featurePoints: epic.featurePoints || 0,
      pointsValue: (epic.featurePoints || 0) * 10,
      storyCount: storyCount,
      storyPoints: totalStoryPoints,
      sizedPercent: sizedPercent,
      allocation: epic.allocation || '',
      sprint: epic.sprintName || '',
      fixVersion: epic.fixVersion || '',
      status: epic.status || '',
      iterationStart: epic.iterationStart || '',
      iterationEnd: epic.iterationEnd || '',
      url: epic.url
    });
  });
  
  // Sort by Cost of Delay (descending)
  epicData.sort((a, b) => b.costOfDelay - a.costOfDelay);
  
  // Write data rows
  epicData.forEach((epic, index) => {
    const rowNum = startRow + index;
    
    // Key (with hyperlink) - JIRA_CONFIG.baseUrl is defined in the main config file
    const epicKeyFormula = `=HYPERLINK("${JIRA_CONFIG.baseUrl}/browse/${epic.key}","${epic.key}")`;
    sheet.getRange(rowNum, 1).setFormula(epicKeyFormula);
    
    // Summary (column B)
    sheet.getRange(rowNum, 2).setValue(epic.summary);
    sheet.getRange(rowNum, 2).setWrap(true);
    
    // Cost of Delay (column C) - format as number with thousands separator
    sheet.getRange(rowNum, 3).setValue(epic.costOfDelay);
    sheet.getRange(rowNum, 3).setNumberFormat('#,##0');
    
    // Feature Points (column D)
    sheet.getRange(rowNum, 4).setValue(epic.featurePoints);
    
    // Feature Point Conversion Value (column E)
    sheet.getRange(rowNum, 5).setValue(epic.pointsValue);
    
    // Ticket Count (column F)
    sheet.getRange(rowNum, 6).setValue(epic.storyCount);
    
    // Story Points (column G)
    sheet.getRange(rowNum, 7).setValue(epic.storyPoints);
    
    // Sized % (column H) - FIXED: Center aligned
    sheet.getRange(rowNum, 8).setValue(epic.sizedPercent + '%');
    sheet.getRange(rowNum, 8).setHorizontalAlignment('center');  // Center align
    
    // Allocation (column I)
    sheet.getRange(rowNum, 9).setValue(epic.allocation);
    
    // Sprint (column J)
    sheet.getRange(rowNum, 10).setValue(epic.sprint);
    
    // Fix Version (column K)
    sheet.getRange(rowNum, 11).setValue(epic.fixVersion);
    
    // Status (column L)
    sheet.getRange(rowNum, 12).setValue(epic.status);
    
    // Iteration Start (column M)
    sheet.getRange(rowNum, 13).setValue(epic.iterationStart);
    
    // Iteration End (column N)
    sheet.getRange(rowNum, 14).setValue(epic.iterationEnd);
    
    // Apply conditional formatting for Sized %
    if (epic.sizedPercent === 100) {
      sheet.getRange(rowNum, 8).setBackground('#ccffcc'); // Green
    } else if (epic.sizedPercent >= 50) {
      sheet.getRange(rowNum, 8).setBackground('#ffffcc'); // Yellow
    } else {
      sheet.getRange(rowNum, 8).setBackground('#ffcccc'); // Red
    }
  });
  
  // Format data rows
  if (epicData.length > 0) {
    sheet.getRange(startRow, 1, epicData.length, headers.length).setFontSize(8).setWrap(true).setFontFamily('Comfortaa').setVerticalAlignment('middle');
    // FIXED: Columns C-G centered (not including H anymore as it's set individually above)
    sheet.getRange(startRow, 3, epicData.length, 5).setHorizontalAlignment('center'); 
    
    // Set reasonable row heights for data rows
    for (let i = 0; i < epicData.length; i++) {
      setRowHeightWithLimit(sheet, startRow + i, 30, 70);
    }
  }
  
  // NO TOTALS ROW as requested - just return the next row after the data
  startRow += epicData.length;
  startRow += 2; // Add spacing after section
  
  return startRow;
}

/**
 * Creates Release Version Validation section for scrum team summary
 * @param {Sheet} sheet - The sheet to write to
 * @param {number} startRow - The row to start writing at
 * @param {Array} issues - All issues for the team
 * @param {string} scrumTeam - The scrum team name
 * @param {string} programIncrement - The PI name (e.g., "PI 12")
 * @return {number} The next available row after the section
 */
function createReleaseVersionValidation(sheet, startRow, issues, scrumTeam, programIncrement) {
  console.log(`Creating Release Version Validation for ${scrumTeam} - checking for epics with iteration dates and children with sprints`);
  
  // Extract PI number from programIncrement (e.g., "PI 12" -> 12)
  const piNumber = parseInt(programIncrement.replace('PI ', ''));
  if (isNaN(piNumber)) {
    console.error('Invalid PI number in programIncrement:', programIncrement);
    return startRow;
  }
  
  // Filter epics and child tickets
  const epics = issues.filter(i => i.issueType === 'Epic');
  const childTickets = issues.filter(i => 
    (i.issueType === 'Story' || i.issueType === 'Bug') && 
    (i.epicLink || i.parentKey)
  );
  
  console.log(`Found ${epics.length} epics and ${childTickets.length} child tickets`);
  
  if (epics.length === 0) {
    console.log('No epics found for Release Version Validation');
    // Add a small note indicating no epics
    sheet.getRange(startRow, 1).setValue('Release Version Validation: No epics found');
    sheet.getRange(startRow, 1).setFontSize(8).setFontStyle('italic').setFontColor('#999999').setFontFamily('Comfortaa');
    setRowHeightWithLimit(sheet, startRow, 20, 70);
    return startRow + 2;
  }
  
  // Process each epic
  const validationData = [];
  let epicsChecked = 0;
  let epicsWithoutDates = 0;
  let epicsWithoutSprints = 0;
  
  epics.forEach(epic => {
    epicsChecked++;
    
    // Skip epics without iteration dates
    if (!epic.iterationStart || !epic.iterationEnd) {
      epicsWithoutDates++;
      console.log(`Skipping epic ${epic.key} - missing iteration dates (Start: ${epic.iterationStart}, End: ${epic.iterationEnd})`);
      return;
    }
    
    // Find all children of this epic
    const epicChildren = childTickets.filter(child => 
      child.epicLink === epic.key || child.parentKey === epic.key
    );
    
    console.log(`Epic ${epic.key} has ${epicChildren.length} children`);
    
    // Check if at least one child has a sprint
    const hasChildWithSprint = epicChildren.some(child => child.sprintName && child.sprintName.trim() !== '');
    
    if (!hasChildWithSprint) {
      epicsWithoutSprints++;
      console.log(`Skipping epic ${epic.key} - no children with sprint assigned`);
      return;
    }
    
    // Determine recommended fix version based on children's sprints
    const recommendedFixVersion = determineRecommendedFixVersion(epicChildren, piNumber);
    
    // Check if there's a mismatch
    const actualFixVersion = epic.fixVersion || '';
    const hasMismatch = recommendedFixVersion && actualFixVersion !== recommendedFixVersion;
    
    // Find children with wrong fix version
    const childrenWithWrongFixVersion = findChildrenWithWrongFixVersion(
      epicChildren, 
      actualFixVersion || recommendedFixVersion
    );
    
    validationData.push({
      epicKey: epic.key,
      summary: epic.summary || '',
      iterationStart: epic.iterationStart || '',
      iterationEnd: epic.iterationEnd || '',
      fixVersion: actualFixVersion,
      recommendedFixVersion: recommendedFixVersion,
      hasMismatch: hasMismatch,
      childrenWithWrongFixVersion: childrenWithWrongFixVersion
    });
  });
  
  console.log(`Release Version Validation Summary: ${epicsChecked} epics checked, ${epicsWithoutDates} without dates, ${epicsWithoutSprints} without sprints, ${validationData.length} with potential issues`);
  
  // If no epics meet the criteria, add a note
  if (validationData.length === 0) {
    console.log('No epics with iteration dates and children with sprints found for Release Version Validation');
    
    // Add a small note indicating no issues found
    let message = 'Release Version Validation: No issues detected';
    if (epicsWithoutDates > 0 || epicsWithoutSprints > 0) {
      message += ` (${epicsWithoutDates} epics without dates, ${epicsWithoutSprints} without sprints)`;
    }
    sheet.getRange(startRow, 1).setValue(message);
    sheet.getRange(startRow, 1).setFontSize(8).setFontStyle('italic').setFontColor('#999999').setFontFamily('Comfortaa');
    setRowHeightWithLimit(sheet, startRow, 20, 70);
    
    return startRow + 2;
  }
  
  // Section title - no merge, fill A, B, C with purple
  sheet.getRange(startRow, 1).setValue('Release Version Validation');
  sheet.getRange(startRow, 1, 1, 3).setBackground('#E1D5E7');
  sheet.getRange(startRow, 1).setFontSize(14).setFontWeight('bold').setFontColor('black');
  sheet.getRange(startRow, 1).setFontFamily('Comfortaa');
  sheet.getRange(startRow, 1).setVerticalAlignment('middle');
  setRowHeightWithLimit(sheet, startRow, 30, 70);
  startRow++; // No space between title and table
  
  // Headers
  const headers = [
    'Epic Key', 'Summary', 'Iteration Start', 'Iteration End', 
    'Fix Version', 'Epic Recommended Fix Version', 'Children with Potential Wrong Fix Version'
  ];
  
  sheet.getRange(startRow, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(startRow, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#9b7bb8')
    .setFontColor('white')
    .setFontSize(8)
    .setWrap(true)
    .setFontFamily('Comfortaa')
    .setVerticalAlignment('middle');
  
  // Set header row height (max 70)
  setRowHeightWithLimit(sheet, startRow, 40, 70);
  
  // Set all column widths to 100
  for (let col = 1; col <= headers.length; col++) {
    sheet.setColumnWidth(col, 100);
  }
  
  startRow++;
  
  // Write data rows
  validationData.forEach((data, index) => {
    const rowNum = startRow + index;
    
    // Epic Key (with hyperlink) - JIRA_CONFIG.baseUrl is defined in the main config file
    const epicKeyFormula = `=HYPERLINK("${JIRA_CONFIG.baseUrl}/browse/${data.epicKey}","${data.epicKey}")`;
    sheet.getRange(rowNum, 1).setFormula(epicKeyFormula);
    
    // Summary
    sheet.getRange(rowNum, 2).setValue(data.summary);
    sheet.getRange(rowNum, 2).setWrap(true);
    
    // Iteration dates
    sheet.getRange(rowNum, 3).setValue(data.iterationStart);
    sheet.getRange(rowNum, 4).setValue(data.iterationEnd);
    
    // Fix Version
    sheet.getRange(rowNum, 5).setValue(data.fixVersion);
    
    // Epic Recommended Fix Version
    sheet.getRange(rowNum, 6).setValue(data.recommendedFixVersion);
    
    // Highlight if mismatch
    if (data.hasMismatch) {
      sheet.getRange(rowNum, 5, 1, 2).setBackground('#ffcccc'); // Light red for both columns
    }
    
    // Children with wrong fix version (as hyperlinks) - JIRA_CONFIG.baseUrl is defined in the main config file
    if (data.childrenWithWrongFixVersion.length > 0) {
      const childLinks = data.childrenWithWrongFixVersion.map(childKey => 
        `=HYPERLINK("${JIRA_CONFIG.baseUrl}/browse/${childKey}","${childKey}")`
      );
      
      // Concatenate formulas with CHAR(10) for line breaks
      const formula = childLinks.length === 1 ? childLinks[0] : 
        `=CONCATENATE(${childLinks.join(',CHAR(10),')})`;
      
      sheet.getRange(rowNum, 7).setFormula(formula);
      sheet.getRange(rowNum, 7).setWrap(true);
    }
    
    // Set row height if needed (capped at 70)
    if (data.childrenWithWrongFixVersion.length > 2) {
      const desiredHeight = 20 + (data.childrenWithWrongFixVersion.length * 15);
      setRowHeightWithLimit(sheet, rowNum, desiredHeight, 70);
    }
  });
  
  // Format data rows
  if (validationData.length > 0) {
    sheet.getRange(startRow, 1, validationData.length, headers.length).setFontSize(8).setWrap(true).setFontFamily('Comfortaa').setVerticalAlignment('middle');
    sheet.getRange(startRow, 1, validationData.length, headers.length).setVerticalAlignment('top');
    
    // Set standard row heights for data rows (unless overridden by children count)
    for (let i = 0; i < validationData.length; i++) {
      const data = validationData[i];
      if (data.childrenWithWrongFixVersion.length <= 2) {
        setRowHeightWithLimit(sheet, startRow + i, 30, 70);
      }
    }
  }
  
  startRow += validationData.length;
  startRow += 2; // Add spacing after section
  
  return startRow;
}

/**
 * Determines the recommended fix version based on children's sprint iterations
 * @param {Array} children - Array of child issues (Stories and Bugs)
 * @param {number} piNumber - The PI number (e.g., 12)
 * @return {string} The recommended fix version or empty string
 */
function determineRecommendedFixVersion(children, piNumber) {
  if (!children || children.length === 0) return '';
  
  let maxIteration = 0;
  let foundValidSprint = false;
  
  // Analyze each child's sprint name
  children.forEach(child => {
    if (child.sprintName) {
      // Extract iteration from sprint name
      // Try multiple patterns: "PI 12.3", "12.3", "Team - 12.3", etc.
      const patterns = [
        new RegExp(`PI\\s*${piNumber}\\s*\\.\\s*(\\d+)`, 'i'),
        new RegExp(`${piNumber}\\s*\\.\\s*(\\d+)`, 'i'),
        new RegExp(`\\b${piNumber}\\.(\\d+)\\b`)
      ];
      
      for (const pattern of patterns) {
        const match = child.sprintName.match(pattern);
        if (match) {
          const sprintIteration = parseInt(match[1]);
          if (sprintIteration > maxIteration) {
            maxIteration = sprintIteration;
            foundValidSprint = true;
          }
          break; // Found a match, no need to try other patterns
        }
      }
    }
  });
  
  if (!foundValidSprint) {
    console.log(`No valid sprint patterns found for PI ${piNumber} in children`);
    return '';
  }
  
  // Apply business rules for fix version
  if (maxIteration >= 5) {
    // If any sprint is in iteration 5 or 6
    return 'Release 7.13';
  } else if (maxIteration >= 1 && maxIteration <= 4) {
    // If all sprints are in iterations 1-4
    return 'Release 7.12';
  }
  
  return '';
}

/**
 * Finds children with wrong or missing fix version
 * @param {Array} children - Array of child issues
 * @param {string} expectedFixVersion - The expected fix version from the epic
 * @return {Array} Array of child keys with wrong fix version
 */
function findChildrenWithWrongFixVersion(children, expectedFixVersion) {
  const wrongChildren = [];
  
  if (!expectedFixVersion) return wrongChildren;
  
  // Normalize expected fix version for comparison
  const expectedNorm = expectedFixVersion.trim().toLowerCase();
  
  children.forEach(child => {
    const childFixVersion = child.fixVersion || '';
    const childNorm = childFixVersion.trim().toLowerCase();
    
    // Check if child has no fix version or different fix version
    if (!childFixVersion || childNorm !== expectedNorm) {
      wrongChildren.push(child.key);
    }
  });
  
  return wrongChildren;
}

/**
 * Integration point - add this call to your createScrumTeamSummary function
 * This should be added after the existing sections but before the end of the summary
 */
function addAllocationMismatchToSummary(sheet, currentRow, teamIssues, scrumTeam) {
  console.log(`Checking allocation mismatches for ${scrumTeam}`);
  currentRow = createEpicAllocationMismatchTable(sheet, currentRow, teamIssues, scrumTeam);
  return currentRow;
}

/**
 * Creates a table showing epics with child tickets that have mismatched allocations
 * @param {Sheet} sheet - The sheet to write to
 * @param {number} startRow - The row to start writing at
 * @param {Array} issues - All issues for the team
 * @param {string} scrumTeam - The scrum team name
 * @return {number} The next available row after the table
 */
function createEpicAllocationMismatchTable(sheet, startRow, issues, scrumTeam) {
  console.log(`Creating allocation mismatch table for ${scrumTeam}`);
  
  // Filter epics and potential child tickets
  const epics = issues.filter(i => i.issueType === 'Epic');
  const childTickets = issues.filter(i => 
    (i.issueType === 'Story' || i.issueType === 'Bug') && 
    (i.epicLink || i.parentKey)
  );
  
  console.log(`Found ${epics.length} epics and ${childTickets.length} child tickets for allocation mismatch check`);
  
  // Find epics with mismatched children
  const epicsWithMismatches = [];
  
  epics.forEach(epic => {
    // Find all children of this epic
    const epicChildren = childTickets.filter(child => 
      child.epicLink === epic.key || child.parentKey === epic.key
    );
    
    // Find children with different allocation than the epic
    const mismatchedChildren = epicChildren.filter(child => {
      // Skip if either allocation is missing
      if (!child.allocation || !epic.allocation) {
        return false;
      }
      
      // Normalize allocations for comparison (trim whitespace, compare case-insensitive)
      const childAllocationNorm = child.allocation.trim().toLowerCase();
      const epicAllocationNorm = epic.allocation.trim().toLowerCase();
      
      // Log the allocation comparison for debugging
      if (epicChildren.indexOf(child) === 0) { // Only log first child
        console.log(`Comparing allocations - Epic ${epic.key}: "${epic.allocation}" vs Child ${child.key}: "${child.allocation}"`);
      }
      
      return childAllocationNorm !== epicAllocationNorm;
    });
    
    if (mismatchedChildren.length > 0) {
      console.log(`Epic ${epic.key} has ${mismatchedChildren.length} mismatched children`);
      epicsWithMismatches.push({
        epic: epic,
        mismatchedChildren: mismatchedChildren
      });
    }
  });
  
  // If no mismatches found, add a note
  if (epicsWithMismatches.length === 0) {
    console.log('No allocation mismatches found');
    
    // Add a small note indicating no issues found
    sheet.getRange(startRow, 1).setValue('Epic-Child Allocation Mismatches: No issues detected');
    sheet.getRange(startRow, 1).setFontSize(8).setFontStyle('italic').setFontColor('#999999').setFontFamily('Comfortaa');
    setRowHeightWithLimit(sheet, startRow, 20, 70);
    
    return startRow + 2;
  }
  
  // Create section title - merge A, B, C
  sheet.getRange(startRow, 1, 1, 3).merge();
  sheet.getRange(startRow, 1).setValue('Epic-Child Allocation Mismatches');
  sheet.getRange(startRow, 1).setFontSize(14).setFontWeight('bold').setBackground('#E1D5E7').setFontColor('black');
  sheet.getRange(startRow, 1).setFontFamily('Comfortaa');
  sheet.getRange(startRow, 1).setVerticalAlignment('middle');
  setRowHeightWithLimit(sheet, startRow, 30, 70);
  startRow++; // No space between title and table
  
  // Create headers
  const headers = ['Epic Key', 'Epic Summary', 'Epic Allocation', 'Mismatched Children'];
  sheet.getRange(startRow, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(startRow, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#9b7bb8')
    .setFontColor('white')
    .setFontSize(8)
    .setWrap(true)
    .setFontFamily('Comfortaa')
    .setVerticalAlignment('middle');
  
  // Set header row height
  setRowHeightWithLimit(sheet, startRow, 30, 70);
  
  // Set all column widths to 100
  for (let col = 1; col <= headers.length; col++) {
    sheet.setColumnWidth(col, 100);
  }
  
  startRow++;
  
  // Process each epic with mismatches
  epicsWithMismatches.forEach(item => {
    const epic = item.epic;
    const children = item.mismatchedChildren;
    
    // Determine how many columns we need for children
    const maxChildrenPerRow = 10; // Limit to prevent excessive horizontal spread
    const numRows = Math.ceil(children.length / maxChildrenPerRow);
    
    for (let rowIndex = 0; rowIndex < numRows; rowIndex++) {
      const rowStartIndex = rowIndex * maxChildrenPerRow;
      const rowEndIndex = Math.min(rowStartIndex + maxChildrenPerRow, children.length);
      const childrenForRow = children.slice(rowStartIndex, rowEndIndex);
      
      // Write epic info (only on first row for each epic)
      if (rowIndex === 0) {
        // Epic Key with hyperlink
        if (epic.url) {
          const richText = SpreadsheetApp.newRichTextValue()
            .setText(epic.key)
            .setLinkUrl(epic.url)
            .build();
          sheet.getRange(startRow, 1).setRichTextValue(richText);
        } else {
          sheet.getRange(startRow, 1).setValue(epic.key);
        }
        
        // Epic Summary
        sheet.getRange(startRow, 2).setValue(epic.summary || '');
        sheet.getRange(startRow, 2).setWrap(true);
        
        // Epic Allocation
        sheet.getRange(startRow, 3).setValue(epic.allocation || 'None');
      }
      
      // Write child tickets info starting from column 4
      let childColumn = 4;
      childrenForRow.forEach(child => {
        // Create child info text with allocation
        const childText = `${child.key}\n(${child.allocation || 'None'})`;
        
        // Set as hyperlink if URL exists
        if (child.url) {
          const richText = SpreadsheetApp.newRichTextValue()
            .setText(childText)
            .setLinkUrl(child.url)
            .build();
          sheet.getRange(startRow, childColumn).setRichTextValue(richText);
        } else {
          sheet.getRange(startRow, childColumn).setValue(childText);
        }
        
        // Style the cell
        sheet.getRange(startRow, childColumn)
          .setWrap(true)
          .setFontSize(8)
          .setVerticalAlignment('top')
          .setBackground('#ffe6e6') // Light red background for mismatches
          .setFontFamily('Comfortaa');
        
        childColumn++;
      });
      
      // Style the row
      sheet.getRange(startRow, 1, 1, 3).setFontSize(8).setVerticalAlignment('top').setWrap(true).setFontFamily('Comfortaa');
      
      // Set row height to accommodate wrapped text (max 70)
      setRowHeightWithLimit(sheet, startRow, 60, 70);
      
      startRow++;
    }
    
    // Add a small gap between epics
    if (epicsWithMismatches.indexOf(item) < epicsWithMismatches.length - 1) {
      sheet.setRowHeight(startRow, 5);
      startRow++;
    }
  });
  
  // Add spacing after the table
  startRow += 2;
  
  return startRow;
}

// ===== UI ENTRY POINTS =====
/**
 * Generate summaries for all scrum teams in a PI
 */
function generateAllScrumTeamSummaries(piNumber) {
  const ui = SpreadsheetApp.getUi();
  const programIncrement = `PI ${piNumber}`;
  
  try {
    showProgress(`Generating summaries for all teams in ${programIncrement}...`);
    
    // Check if PI data sheet exists
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const piSheetName = `PI ${piNumber}`;
    const piSheet = spreadsheet.getSheetByName(piSheetName);
    
    if (!piSheet) {
      closeProgress();
      ui.alert(`No data found for ${programIncrement}. Please run the full analysis first.`);
      return;
    }
    
    // Read data from the PI sheet
    showProgress('Reading PI data...');
    const dataRange = piSheet.getDataRange();
    const values = dataRange.getValues();
    const headers = values[3];
    
    const allIssues = parsePISheetData(values, headers);
    
    // Get unique scrum teams
    const scrumTeams = [...new Set(allIssues.map(issue => issue.scrumTeam || 'Unassigned'))].sort();
    
    if (scrumTeams.length === 0) {
      closeProgress();
      ui.alert('No scrum teams found in the data.');
      return;
    }
    
    // Generate summaries for all teams
    showProgress(`Creating summaries for ${scrumTeams.length} teams...`);
    createScrumTeamSummaries(allIssues, programIncrement, scrumTeams);
    
    closeProgress();
    
    ui.alert(
      'Success',
      `Summary reports generated for ${scrumTeams.length} teams in ${programIncrement}:\n\n` +
      scrumTeams.join('\n'),
      ui.ButtonSet.OK
    );
    
  } catch (error) {
    console.error('Error generating summaries:', error);
    closeProgress();
    ui.alert('Error', 'An error occurred: ' + error.toString(), ui.ButtonSet.OK);
  }
}

/**
 * Batch update function for multiple teams
 * This function can be called to update multiple team summaries at once
 * @param {string} piNumber - The PI number
 * @param {Array} teamsToUpdate - Array of team names to update
 * @return {Object} Results object with success/failure counts
 */
function batchUpdateScrumTeamSummaries(piNumber, teamsToUpdate) {
  const programIncrement = `PI ${piNumber}`;
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const piSheet = spreadsheet.getSheetByName(`PI ${piNumber}`);
  
  if (!piSheet) {
    throw new Error(`No data found for ${programIncrement}. Please run the full analysis first.`);
  }
  
  // Read all data from the PI sheet
  const dataRange = piSheet.getDataRange();
  const values = dataRange.getValues();
  const headers = values[3];
  const allIssues = parsePISheetData(values, headers);
  
  const results = {
    success: [],
    failed: [],
    noData: [],
    total: teamsToUpdate.length
  };
  
  // Process each team
  teamsToUpdate.forEach(team => {
    try {
      showProgress(`Updating summary for ${team}...`);
      
      const result = createScrumTeamSummary(allIssues, programIncrement, team);
      
      if (result.success) {
        results.success.push(team);
      } else if (result.error && result.error.includes('No data found')) {
        results.noData.push({
          team: team,
          error: result.error
        });
      } else {
        results.failed.push({
          team: team,
          error: result.error || 'Unknown error'
        });
      }
      
    } catch (error) {
      console.error(`Failed to update ${team}:`, error);
      results.failed.push({
        team: team,
        error: error.toString()
      });
    }
  });
  
  return results;
}

/**
 * Creates a menu in the Google Sheets UI for easy access to debug functions
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Scrum Summary Tools')
    .addItem('Debug Team Data...', 'showDebugDialog')
    .addToUi();
}

/**
 * Shows a dialog to input debug parameters
 */
function showDebugDialog() {
  const html = HtmlService.createHtmlOutputFromString(`
    <div style="padding: 10px;">
      <label for="piNumber">PI Number:</label><br>
      <input type="number" id="piNumber" value="12" style="margin: 5px 0;"><br>
      <label for="teamName">Team Name:</label><br>
      <input type="text" id="teamName" value="ATLAS" style="margin: 5px 0;"><br>
      <button onclick="runDebug()">Run Debug</button>
    </div>
    <script>
      function runDebug() {
        const piNumber = document.getElementById('piNumber').value;
        const teamName = document.getElementById('teamName').value;
        google.script.run.debugScrumTeamData(piNumber, teamName);
        google.script.host.close();
      }
    </script>
  `)
  .setWidth(300)
  .setHeight(200);
  
  SpreadsheetApp.getUi().showModalDialog(html, 'Debug Team Data');
}

/**
 * Debug function to check parsed data for a specific team
 * @param {string} piNumber - The PI number
 * @param {string} scrumTeamName - The scrum team name
 */
function debugScrumTeamData(piNumber, scrumTeamName) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const piSheet = spreadsheet.getSheetByName(`PI ${piNumber}`);
  
  if (!piSheet) {
    console.log(`No sheet found for PI ${piNumber}`);
    return;
  }
  
  const dataRange = piSheet.getDataRange();
  const values = dataRange.getValues();
  const headers = values[3];
  
  console.log('Headers found:', headers);
  
  const allIssues = parsePISheetData(values, headers);
  const teamIssues = allIssues.filter(issue => 
    (issue.scrumTeam || 'Unassigned') === scrumTeamName
  );
  
  console.log(`Found ${teamIssues.length} issues for team ${scrumTeamName}`);
  
  // Check for epics with iteration dates
  const epics = teamIssues.filter(i => i.issueType === 'Epic');
  console.log(`Found ${epics.length} epics`);
  
  const epicsWithDates = epics.filter(e => e.iterationStart && e.iterationEnd);
  console.log(`Found ${epicsWithDates.length} epics with iteration dates`);
  
  if (epicsWithDates.length > 0) {
    console.log('Sample epic with dates:', {
      key: epicsWithDates[0].key,
      iterationStart: epicsWithDates[0].iterationStart,
      iterationEnd: epicsWithDates[0].iterationEnd,
      fixVersion: epicsWithDates[0].fixVersion
    });
  }
  
  // Check for allocation mismatches
  const epicKeys = epics.map(e => e.key);
  const childTickets = teamIssues.filter(i => 
    (i.issueType === 'Story' || i.issueType === 'Bug') && 
    (epicKeys.includes(i.epicLink) || epicKeys.includes(i.parentKey))
  );
  
  console.log(`Found ${childTickets.length} child tickets of epics`);
  
  // Check allocations
  const sampleEpic = epics[0];
  if (sampleEpic) {
    console.log(`Sample epic allocation: ${sampleEpic.key} = "${sampleEpic.allocation}"`);
    const epicChildren = childTickets.filter(c => 
      c.epicLink === sampleEpic.key || c.parentKey === sampleEpic.key
    );
    if (epicChildren.length > 0) {
      console.log(`Sample child allocation: ${epicChildren[0].key} = "${epicChildren[0].allocation}"`);
    }
  }
}

/**
 * Creates a more detailed allocation mismatch report with better formatting
 * Alternative implementation with a different layout
 */
function createDetailedAllocationMismatchReport(sheet, startRow, issues, scrumTeam) {
  console.log(`Creating detailed allocation mismatch report for ${scrumTeam}`);
  
  const epics = issues.filter(i => i.issueType === 'Epic');
  const childTickets = issues.filter(i => 
    (i.issueType === 'Story' || i.issueType === 'Bug') && 
    (i.epicLink || i.parentKey)
  );
  
  // Collect all mismatches
  const allMismatches = [];
  
  epics.forEach(epic => {
    const epicChildren = childTickets.filter(child => 
      child.epicLink === epic.key || child.parentKey === epic.key
    );
    
    epicChildren.forEach(child => {
      if (child.allocation && epic.allocation && child.allocation !== epic.allocation) {
        allMismatches.push({
          epicKey: epic.key,
          epicSummary: epic.summary || '',
          epicAllocation: epic.allocation,
          epicUrl: epic.url,
          childKey: child.key,
          childSummary: child.summary || '',
          childAllocation: child.allocation,
          childType: child.issueType,
          childUrl: child.url,
          childStoryPoints: child.storyPoints || 0
        });
      }
    });
  });
  
  if (allMismatches.length === 0) {
    console.log('No allocation mismatches found');
    return startRow;
  }
  
  // Section title - merge A, B, C
  sheet.getRange(startRow, 1, 1, 3).merge();
  sheet.getRange(startRow, 1).setValue('Epic-Child Allocation Mismatch Details');
  sheet.getRange(startRow, 1).setFontSize(14).setFontWeight('bold').setBackground('#E1D5E7').setFontColor('black');
  sheet.getRange(startRow, 1).setFontFamily('Comfortaa');
  sheet.getRange(startRow, 1).setVerticalAlignment('middle');
  setRowHeightWithLimit(sheet, startRow, 30, 70);
  startRow++; // No space between title and table
  
  // Summary info
  const uniqueEpics = new Set(allMismatches.map(m => m.epicKey)).size;
  sheet.getRange(startRow, 1).setValue(`Found ${allMismatches.length} mismatched child tickets across ${uniqueEpics} epics`);
  sheet.getRange(startRow, 1).setFontStyle('italic').setFontSize(8).setFontFamily('Comfortaa');
  setRowHeightWithLimit(sheet, startRow, 20, 70);
  startRow += 2;
  
  // Table headers
  const headers = [
    'Epic Key', 'Epic Summary', 'Epic Allocation', 
    'Child Key', 'Child Type', 'Child Allocation', 'Story Points'
  ];
  
  sheet.getRange(startRow, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(startRow, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#9b7bb8')
    .setFontColor('white')
    .setFontSize(8)
    .setWrap(true)
    .setFontFamily('Comfortaa')
    .setVerticalAlignment('middle');
  
  // Set header row height (max 70)
  setRowHeightWithLimit(sheet, startRow, 30, 70);
  
  // Set all column widths to 100
  for (let col = 1; col <= headers.length; col++) {
    sheet.setColumnWidth(col, 100);
  }
  
  startRow++;
  
  // Sort mismatches by epic key, then child key
  allMismatches.sort((a, b) => {
    if (a.epicKey !== b.epicKey) return a.epicKey.localeCompare(b.epicKey);
    return a.childKey.localeCompare(b.childKey);
  });
  
  // Write data rows
  let currentEpic = null;
  allMismatches.forEach((mismatch, index) => {
    const isNewEpic = mismatch.epicKey !== currentEpic;
    currentEpic = mismatch.epicKey;
    
    // Add visual separator for new epics
    if (isNewEpic && index > 0) {
      sheet.getRange(startRow, 1, 1, headers.length).setBorder(true, false, false, false, false, false, 'black', SpreadsheetApp.BorderStyle.SOLID);
    }
    
    // Epic Key (with hyperlink)
    if (mismatch.epicUrl) {
      const richText = SpreadsheetApp.newRichTextValue()
        .setText(mismatch.epicKey)
        .setLinkUrl(mismatch.epicUrl)
        .build();
      sheet.getRange(startRow, 1).setRichTextValue(richText);
    } else {
      sheet.getRange(startRow, 1).setValue(mismatch.epicKey);
    }
    
    // Epic Summary (only show for first row of each epic)
    if (isNewEpic) {
      sheet.getRange(startRow, 2).setValue(mismatch.epicSummary);
      sheet.getRange(startRow, 2).setWrap(true);
      sheet.getRange(startRow, 3).setValue(mismatch.epicAllocation);
    }
    
    // Child Key (with hyperlink)
    if (mismatch.childUrl) {
      const richText = SpreadsheetApp.newRichTextValue()
        .setText(mismatch.childKey)
        .setLinkUrl(mismatch.childUrl)
        .build();
      sheet.getRange(startRow, 4).setRichTextValue(richText);
    } else {
      sheet.getRange(startRow, 4).setValue(mismatch.childKey);
    }
    
    // Child details
    sheet.getRange(startRow, 5).setValue(mismatch.childType);
    sheet.getRange(startRow, 6).setValue(mismatch.childAllocation);
    sheet.getRange(startRow, 7).setValue(mismatch.childStoryPoints);
    
    // Highlight the allocation mismatch cells
    sheet.getRange(startRow, 3).setBackground('#ffeeee'); // Epic allocation
    sheet.getRange(startRow, 6).setBackground('#ffcccc'); // Child allocation (darker)
    
    // Style the row - changed font size to 8
    sheet.getRange(startRow, 1, 1, headers.length).setFontSize(8).setWrap(true).setFontFamily('Comfortaa').setVerticalAlignment('middle');
    
    // Set standard row height
    setRowHeightWithLimit(sheet, startRow, 25, 70);
    
    startRow++;
  });
  
  // Add totals row
  startRow++;
  sheet.getRange(startRow, 1, 1, 6).merge();
  sheet.getRange(startRow, 1).setValue('Total Story Points in Mismatched Items:');
  sheet.getRange(startRow, 1).setHorizontalAlignment('right').setFontWeight('bold').setFontSize(8).setFontFamily('Comfortaa').setVerticalAlignment('middle');
  
  const totalPoints = allMismatches.reduce((sum, m) => sum + m.childStoryPoints, 0);
  sheet.getRange(startRow, 7).setValue(totalPoints);
  sheet.getRange(startRow, 7).setFontWeight('bold').setFontSize(8).setFontFamily('Comfortaa').setVerticalAlignment('middle');
  
  // Set row height for totals row
  setRowHeightWithLimit(sheet, startRow, 25, 70);
  
  startRow += 3;
  
  return startRow;
}
