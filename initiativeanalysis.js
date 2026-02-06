/**
 * Initiative Analysis - Creates Initiative Analysis tabs per value stream
 * 
 * This module generates summary tabs showing Portfolio Initiative and Program Initiative
 * distribution for each value stream, including pie charts for visualization.
 * 
 * Usage:
 * - Call generateInitiativeAnalysisForValueStream(piNumber, valueStream) for a single value stream
 * - Call generateAllInitiativeAnalysisTabs(piNumber) for all value streams
 * - Integrates with existing PI analysis flow
 * 
 * Data Sources:
 * - PI data sheet (e.g., "PI 14") for epic data
 * - Uses portfolioInitiative and programInitiative fields from epics
 * 
 * Calculations:
 * - Feature Points x 10 for point calculations (consistent with existing system)
 * - Falls back to Story Point Estimate if Feature Points not available
 */

// ===== CONFIGURATION =====

const INITIATIVE_ANALYSIS_CONFIG = {
  sheetNamePattern: '{PI} - {VS} Initiatives',  // e.g., "PI 14 - MMPM Initiatives"
  
  // Colors matching existing ModMed theme
  colors: {
    headerPrimary: '#1B365D',      // Navy Blue
    headerSecondary: '#6B3FA0',    // Purple Dark
    sectionHeader: '#9b7bb8',      // Purple Light
    goldAccent: '#FFC72C',         // Gold Yellow
    backgroundLight: '#F5F5F5',
    white: '#FFFFFF',
    purpleLight: '#E1D5E7'
  },
  
  // Chart settings
  chartWidth: 500,
  chartHeight: 350,
  chartColumnOffset: 6,  // Column F for chart placement
  
  // Display settings
  maxInitiativeNameLength: 60,
  fontFamily: 'Comfortaa'
};

// ===== MAIN FUNCTIONS =====

/**
 * Generate Initiative Analysis tab for a specific value stream
 * @param {number|string} piNumber - The PI number (e.g., 14 or "14")
 * @param {string} valueStream - The value stream name (e.g., "MMPM", "EMA Clinical")
 * @param {Spreadsheet} spreadsheet - Optional spreadsheet object (uses active if not provided)
 * @returns {boolean} Success status
 */
function generateInitiativeAnalysisForValueStream(piNumber, valueStream, spreadsheet = null) {
  try {
    const ss = spreadsheet || SpreadsheetApp.getActiveSpreadsheet();
    const programIncrement = `PI ${piNumber}`;
    
    console.log(`Generating Initiative Analysis for ${valueStream} in ${programIncrement}`);
    
    // Get PI data sheet
    const piSheet = ss.getSheetByName(programIncrement);
    if (!piSheet) {
      console.error(`PI sheet "${programIncrement}" not found`);
      return false;
    }
    
    // Read and parse PI data
    const epics = getEpicsForValueStream(piSheet, valueStream);
    
    if (epics.length === 0) {
      console.log(`No epics found for ${valueStream} in ${programIncrement}`);
      return false;
    }
    
    console.log(`Found ${epics.length} epics for ${valueStream}`);
    
    // Create or get the initiative analysis sheet
    const sheetName = INITIATIVE_ANALYSIS_CONFIG.sheetNamePattern
      .replace('{PI}', programIncrement)
      .replace('{VS}', valueStream);
    
    let sheet = ss.getSheetByName(sheetName);
    if (sheet) {
      sheet.clear();
      // Remove existing charts
      const charts = sheet.getCharts();
      charts.forEach(chart => sheet.removeChart(chart));
    } else {
      sheet = ss.insertSheet(sheetName);
    }
    
    // Write the initiative analysis
    writeInitiativeAnalysisSheet(sheet, epics, programIncrement, valueStream);
    
    console.log(`Successfully created ${sheetName}`);
    return true;
    
  } catch (error) {
    console.error(`Error generating Initiative Analysis for ${valueStream}:`, error);
    return false;
  }
}

/**
 * Generate Initiative Analysis tabs for all value streams in a PI
 * @param {number|string} piNumber - The PI number
 * @returns {Object} { success: boolean, created: string[], failed: string[] }
 */
function generateAllInitiativeAnalysisTabs(piNumber) {
  const results = { success: true, created: [], failed: [] };
  
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const programIncrement = `PI ${piNumber}`;
    
    // Get PI data sheet
    const piSheet = spreadsheet.getSheetByName(programIncrement);
    if (!piSheet) {
      console.error(`PI sheet "${programIncrement}" not found`);
      results.success = false;
      return results;
    }
    
    // Get all unique value streams from the PI data
    const valueStreams = getUniqueValueStreamsFromPISheet(piSheet);
    
    console.log(`Found ${valueStreams.length} value streams: ${valueStreams.join(', ')}`);
    
    // Generate initiative analysis for each value stream
    valueStreams.forEach(vs => {
      const success = generateInitiativeAnalysisForValueStream(piNumber, vs, spreadsheet);
      if (success) {
        results.created.push(vs);
      } else {
        results.failed.push(vs);
      }
    });
    
    results.success = results.failed.length === 0;
    
    console.log(`Initiative Analysis complete: ${results.created.length} created, ${results.failed.length} failed`);
    
    return results;
    
  } catch (error) {
    console.error('Error generating all Initiative Analysis tabs:', error);
    results.success = false;
    return results;
  }
}

/**
 * Menu function to generate Initiative Analysis for all value streams
 */
function menuGenerateInitiativeAnalysis() {
  const ui = SpreadsheetApp.getUi();
  
  // Prompt for PI number
  const response = ui.prompt(
    'Generate Initiative Analysis',
    'Enter the PI number (e.g., 14):',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  
  const piNumber = response.getResponseText().trim();
  if (!piNumber || isNaN(parseInt(piNumber))) {
    ui.alert('Invalid PI number. Please enter a valid number.');
    return;
  }
  
  // Show progress
  if (typeof showProgress === 'function') {
    showProgress('Generating Initiative Analysis tabs...');
  }
  
  const results = generateAllInitiativeAnalysisTabs(parseInt(piNumber));
  
  if (typeof closeProgress === 'function') {
    closeProgress();
  }
  
  // Show results
  if (results.success) {
    ui.alert(
      'Initiative Analysis Complete',
      `Successfully created Initiative Analysis tabs for:\n${results.created.join('\n')}`,
      ui.ButtonSet.OK
    );
  } else {
    let message = '';
    if (results.created.length > 0) {
      message += `Created: ${results.created.join(', ')}\n\n`;
    }
    if (results.failed.length > 0) {
      message += `Failed: ${results.failed.join(', ')}`;
    }
    ui.alert('Initiative Analysis Results', message, ui.ButtonSet.OK);
  }
}

// ===== SHEET WRITING FUNCTIONS =====

/**
 * Write Initiative Analysis sheet with pie charts for Portfolio and Program Initiatives
 * @param {Sheet} sheet - The sheet to write to
 * @param {Array} epics - Array of epic objects
 * @param {string} programIncrement - The PI string (e.g., "PI 14")
 * @param {string} valueStream - The value stream name
 */
function writeInitiativeAnalysisSheet(sheet, epics, programIncrement, valueStream) {
  const colors = INITIATIVE_ANALYSIS_CONFIG.colors;
  const fontFamily = INITIATIVE_ANALYSIS_CONFIG.fontFamily;
  let currentRow = 1;
  
  // ===== TITLE SECTION =====
  sheet.getRange(currentRow, 1).setValue(`${programIncrement} - ${valueStream} Initiative Analysis`);
  sheet.getRange(currentRow, 1, 1, 5)
    .setFontSize(18)
    .setFontWeight('bold')
    .setFontColor(colors.white)
    .setBackground(colors.headerPrimary)
    .setFontFamily(fontFamily);
  sheet.getRange(currentRow, 1, 1, 5).merge();
  sheet.setRowHeight(currentRow, 35);
  currentRow++;
  
  // Last Refreshed timestamp
  const now = new Date();
  const formattedDate = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  sheet.getRange(currentRow, 1).setValue(`Last Refreshed: ${formattedDate}`);
  sheet.getRange(currentRow, 1)
    .setFontStyle('italic')
    .setFontColor('#666666')
    .setFontSize(10)
    .setFontFamily(fontFamily);
  currentRow++;
  
  // Summary stats
  const totalEpics = epics.length;
  const totalPoints = epics.reduce((sum, e) => sum + calculateEpicPoints(e), 0);
  sheet.getRange(currentRow, 1).setValue(`Total Epics: ${totalEpics} | Total Points: ${Math.ceil(totalPoints)}`);
  sheet.getRange(currentRow, 1)
    .setFontWeight('bold')
    .setFontSize(11)
    .setFontFamily(fontFamily);
  currentRow += 2;
  
  // ===== PORTFOLIO INITIATIVE SECTION =====
  sheet.getRange(currentRow, 1).setValue('Portfolio Initiative Distribution');
  sheet.getRange(currentRow, 1, 1, 5)
    .setFontSize(14)
    .setFontWeight('bold')
    .setBackground(colors.purpleLight)
    .setFontFamily(fontFamily);
  sheet.getRange(currentRow, 1, 1, 5).merge();
  sheet.setRowHeight(currentRow, 28);
  currentRow += 2;
  
  // Calculate portfolio distribution
  const portfolioData = calculateInitiativeDistribution(epics, 'portfolioInitiative');
  
  // Write portfolio table
  const portfolioHeaders = ['Portfolio Initiative', 'Epic Count', 'Total Points', '% of Total'];
  sheet.getRange(currentRow, 1, 1, portfolioHeaders.length).setValues([portfolioHeaders]);
  sheet.getRange(currentRow, 1, 1, portfolioHeaders.length)
    .setFontWeight('bold')
    .setBackground(colors.headerPrimary)
    .setFontColor(colors.white)
    .setFontSize(10)
    .setFontFamily(fontFamily)
    .setHorizontalAlignment('center');
  sheet.setRowHeight(currentRow, 25);
  currentRow++;
  
  const portfolioStartRow = currentRow;
  if (portfolioData.rows.length > 0) {
    sheet.getRange(currentRow, 1, portfolioData.rows.length, portfolioHeaders.length)
      .setValues(portfolioData.rows);
    sheet.getRange(currentRow, 1, portfolioData.rows.length, portfolioHeaders.length)
      .setFontSize(9)
      .setFontFamily(fontFamily)
      .setVerticalAlignment('middle');
    sheet.getRange(currentRow, 2, portfolioData.rows.length, 3)
      .setHorizontalAlignment('center');
    
    // Alternate row coloring
    for (let i = 0; i < portfolioData.rows.length; i++) {
      if (i % 2 === 1) {
        sheet.getRange(currentRow + i, 1, 1, portfolioHeaders.length)
          .setBackground(colors.backgroundLight);
      }
    }
    
    currentRow += portfolioData.rows.length;
    
    // Totals row
    const portfolioTotals = ['TOTAL', portfolioData.totalEpics, Math.ceil(portfolioData.totalPoints), '100%'];
    sheet.getRange(currentRow, 1, 1, portfolioHeaders.length).setValues([portfolioTotals]);
    sheet.getRange(currentRow, 1, 1, portfolioHeaders.length)
      .setFontWeight('bold')
      .setBackground(colors.goldAccent)
      .setFontSize(10)
      .setFontFamily(fontFamily);
    sheet.getRange(currentRow, 2, 1, 3).setHorizontalAlignment('center');
    currentRow++;
    
    // Create portfolio pie chart
    createInitiativePieChart(
      sheet, 
      portfolioStartRow, 
      portfolioData.rows.length,
      `${valueStream} - Portfolio Initiative Distribution`, 
      INITIATIVE_ANALYSIS_CONFIG.chartColumnOffset, 
      portfolioStartRow - 2
    );
  } else {
    sheet.getRange(currentRow, 1).setValue('No Portfolio Initiative data available');
    sheet.getRange(currentRow, 1).setFontStyle('italic').setFontColor('#666666');
    currentRow++;
  }
  
  currentRow += 3;
  
  // ===== PROGRAM INITIATIVE SECTION =====
  sheet.getRange(currentRow, 1).setValue('Program Initiative Distribution');
  sheet.getRange(currentRow, 1, 1, 5)
    .setFontSize(14)
    .setFontWeight('bold')
    .setBackground(colors.purpleLight)
    .setFontFamily(fontFamily);
  sheet.getRange(currentRow, 1, 1, 5).merge();
  sheet.setRowHeight(currentRow, 28);
  currentRow += 2;
  
  // Calculate program initiative distribution
  const programData = calculateInitiativeDistribution(epics, 'programInitiative');
  
  // Write program initiative table
  const programHeaders = ['Program Initiative', 'Epic Count', 'Total Points', '% of Total'];
  sheet.getRange(currentRow, 1, 1, programHeaders.length).setValues([programHeaders]);
  sheet.getRange(currentRow, 1, 1, programHeaders.length)
    .setFontWeight('bold')
    .setBackground(colors.headerPrimary)
    .setFontColor(colors.white)
    .setFontSize(10)
    .setFontFamily(fontFamily)
    .setHorizontalAlignment('center');
  sheet.setRowHeight(currentRow, 25);
  currentRow++;
  
  const programStartRow = currentRow;
  if (programData.rows.length > 0) {
    sheet.getRange(currentRow, 1, programData.rows.length, programHeaders.length)
      .setValues(programData.rows);
    sheet.getRange(currentRow, 1, programData.rows.length, programHeaders.length)
      .setFontSize(9)
      .setFontFamily(fontFamily)
      .setVerticalAlignment('middle');
    sheet.getRange(currentRow, 2, programData.rows.length, 3)
      .setHorizontalAlignment('center');
    
    // Alternate row coloring
    for (let i = 0; i < programData.rows.length; i++) {
      if (i % 2 === 1) {
        sheet.getRange(currentRow + i, 1, 1, programHeaders.length)
          .setBackground(colors.backgroundLight);
      }
    }
    
    currentRow += programData.rows.length;
    
    // Totals row
    const programTotals = ['TOTAL', programData.totalEpics, Math.ceil(programData.totalPoints), '100%'];
    sheet.getRange(currentRow, 1, 1, programHeaders.length).setValues([programTotals]);
    sheet.getRange(currentRow, 1, 1, programHeaders.length)
      .setFontWeight('bold')
      .setBackground(colors.goldAccent)
      .setFontSize(10)
      .setFontFamily(fontFamily);
    sheet.getRange(currentRow, 2, 1, 3).setHorizontalAlignment('center');
    currentRow++;
    
    // Create program initiative pie chart
    createInitiativePieChart(
      sheet, 
      programStartRow, 
      programData.rows.length,
      `${valueStream} - Program Initiative Distribution`, 
      INITIATIVE_ANALYSIS_CONFIG.chartColumnOffset, 
      programStartRow - 2
    );
  } else {
    sheet.getRange(currentRow, 1).setValue('No Program Initiative data available');
    sheet.getRange(currentRow, 1).setFontStyle('italic').setFontColor('#666666');
    currentRow++;
  }
  
  currentRow += 3;
  
  // ===== ALLOCATION BY INITIATIVE SECTION =====
  currentRow = writeAllocationByInitiativeSection(sheet, currentRow, epics);
  
  currentRow += 2;
  
  // ===== VALUE STREAM CAPACITY DISTRIBUTION =====
  currentRow = writeValueStreamCapacityDistribution(sheet, currentRow, valueStream);
  
  // ===== FORMAT COLUMNS =====
  sheet.setColumnWidth(1, 350);  // Initiative names
  sheet.setColumnWidth(2, 100);  // Epic Count
  sheet.setColumnWidth(3, 100);  // Total Points
  sheet.setColumnWidth(4, 100);  // % of Total
  sheet.setColumnWidth(5, 50);   // Spacer
  
  // Set chart columns wider
  for (let col = 6; col <= 12; col++) {
    sheet.setColumnWidth(col, 80);
  }
  
  console.log(`Created Initiative Analysis sheet with ${portfolioData.rows.length} portfolio and ${programData.rows.length} program initiatives`);
}

/**
 * Write Allocation by Initiative breakdown section
 * Shows how each initiative's points are distributed across allocations
 * @param {Sheet} sheet - The sheet to write to
 * @param {number} startRow - Starting row
 * @param {Array} epics - Array of epic objects
 * @returns {number} Next available row
 */
function writeAllocationByInitiativeSection(sheet, startRow, epics) {
  const colors = INITIATIVE_ANALYSIS_CONFIG.colors;
  const fontFamily = INITIATIVE_ANALYSIS_CONFIG.fontFamily;
  let currentRow = startRow;
  
  // Section header
  sheet.getRange(currentRow, 1).setValue('Allocation Breakdown by Portfolio Initiative');
  sheet.getRange(currentRow, 1, 1, 6)
    .setFontSize(14)
    .setFontWeight('bold')
    .setBackground(colors.purpleLight)
    .setFontFamily(fontFamily);
  sheet.getRange(currentRow, 1, 1, 6).merge();
  sheet.setRowHeight(currentRow, 28);
  currentRow += 2;
  
  // Calculate allocation breakdown per initiative
  const allocationData = calculateAllocationByInitiative(epics, 'portfolioInitiative');
  
  if (allocationData.rows.length === 0) {
    sheet.getRange(currentRow, 1).setValue('No allocation data available');
    sheet.getRange(currentRow, 1).setFontStyle('italic').setFontColor('#666666');
    return currentRow + 2;
  }
  
  // Headers
  const headers = ['Portfolio Initiative', 'Product', 'Tech/Platform', 'Quality', 'KLO', 'Total'];
  sheet.getRange(currentRow, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(currentRow, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground(colors.headerPrimary)
    .setFontColor(colors.white)
    .setFontSize(10)
    .setFontFamily(fontFamily)
    .setHorizontalAlignment('center');
  sheet.setRowHeight(currentRow, 25);
  currentRow++;
  
  // Data rows
  sheet.getRange(currentRow, 1, allocationData.rows.length, headers.length)
    .setValues(allocationData.rows);
  sheet.getRange(currentRow, 1, allocationData.rows.length, headers.length)
    .setFontSize(9)
    .setFontFamily(fontFamily)
    .setVerticalAlignment('middle');
  sheet.getRange(currentRow, 2, allocationData.rows.length, headers.length - 1)
    .setHorizontalAlignment('center');
  
  // Alternate row coloring
  for (let i = 0; i < allocationData.rows.length; i++) {
    if (i % 2 === 1) {
      sheet.getRange(currentRow + i, 1, 1, headers.length)
        .setBackground(colors.backgroundLight);
    }
  }
  
  currentRow += allocationData.rows.length;
  
  // Totals row
  sheet.getRange(currentRow, 1, 1, headers.length).setValues([allocationData.totals]);
  sheet.getRange(currentRow, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground(colors.goldAccent)
    .setFontSize(10)
    .setFontFamily(fontFamily);
  sheet.getRange(currentRow, 2, 1, headers.length - 1).setHorizontalAlignment('center');
  
  return currentRow + 2;
}

// ===== DATA CALCULATION FUNCTIONS =====

/**
 * Calculate initiative distribution for a given field
 * @param {Array} epics - Array of epic objects
 * @param {string} field - Epic field to group by ('portfolioInitiative' or 'programInitiative')
 * @returns {Object} - { rows: [[name, epicCount, points, '%'], ...], totalEpics, totalPoints }
 */
function calculateInitiativeDistribution(epics, field) {
  const distribution = {};
  let totalPoints = 0;
  
  epics.forEach(epic => {
    const initiative = epic[field] || 'Not Specified';
    const estimate = calculateEpicPoints(epic);
    
    if (!distribution[initiative]) {
      distribution[initiative] = { epicCount: 0, points: 0 };
    }
    distribution[initiative].epicCount++;
    distribution[initiative].points += estimate;
    totalPoints += estimate;
  });
  
  // Convert to sorted array (descending by points)
  const maxLen = INITIATIVE_ANALYSIS_CONFIG.maxInitiativeNameLength;
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
 * Calculate allocation breakdown by initiative
 * @param {Array} epics - Array of epic objects
 * @param {string} initiativeField - Field to group by
 * @returns {Object} - { rows: [[name, product, tech, quality, klo, total], ...], totals: [...] }
 */
function calculateAllocationByInitiative(epics, initiativeField) {
  const breakdown = {};
  const allocationTotals = { product: 0, tech: 0, quality: 0, klo: 0 };
  
  epics.forEach(epic => {
    const initiative = epic[initiativeField] || 'Not Specified';
    const points = calculateEpicPoints(epic);
    const allocation = (epic.allocation || '').toLowerCase();
    
    if (!breakdown[initiative]) {
      breakdown[initiative] = { product: 0, tech: 0, quality: 0, klo: 0, total: 0 };
    }
    
    // Categorize allocation
    if (allocation.includes('feature') || allocation.includes('compliance') || allocation.includes('product')) {
      breakdown[initiative].product += points;
      allocationTotals.product += points;
    } else if (allocation.includes('tech') || allocation.includes('platform')) {
      breakdown[initiative].tech += points;
      allocationTotals.tech += points;
    } else if (allocation.includes('quality') || allocation.includes('defect')) {
      breakdown[initiative].quality += points;
      allocationTotals.quality += points;
    } else if (allocation.includes('klo') || allocation.includes('light')) {
      breakdown[initiative].klo += points;
      allocationTotals.klo += points;
    } else {
      // Default to product if no match
      breakdown[initiative].product += points;
      allocationTotals.product += points;
    }
    
    breakdown[initiative].total += points;
  });
  
  // Convert to rows, sorted by total descending
  const maxLen = INITIATIVE_ANALYSIS_CONFIG.maxInitiativeNameLength;
  const rows = Object.entries(breakdown)
    .sort((a, b) => b[1].total - a[1].total)
    .map(([name, data]) => [
      name.length > maxLen ? name.substring(0, maxLen - 3) + '...' : name,
      Math.ceil(data.product),
      Math.ceil(data.tech),
      Math.ceil(data.quality),
      Math.ceil(data.klo),
      Math.ceil(data.total)
    ]);
  
  const grandTotal = allocationTotals.product + allocationTotals.tech + 
                     allocationTotals.quality + allocationTotals.klo;
  
  const totals = [
    'TOTAL',
    Math.ceil(allocationTotals.product),
    Math.ceil(allocationTotals.tech),
    Math.ceil(allocationTotals.quality),
    Math.ceil(allocationTotals.klo),
    Math.ceil(grandTotal)
  ];
  
  return { rows, totals };
}

/**
 * Calculate points for an epic using Feature Points x 10 or Story Point Estimate
 * @param {Object} epic - Epic object
 * @returns {number} Calculated points
 */
function calculateEpicPoints(epic) {
  // Use Feature Points x 10 if available, otherwise Story Point Estimate
  if (epic.featurePoints && epic.featurePoints > 0) {
    return epic.featurePoints * 10;
  }
  return epic.storyPointEstimate || 0;
}

// ===== DATA READING FUNCTIONS =====

/**
 * Get epics for a specific value stream from PI sheet
 * @param {Sheet} piSheet - The PI data sheet
 * @param {string} valueStream - The value stream to filter for
 * @returns {Array} Array of epic objects
 */
function getEpicsForValueStream(piSheet, valueStream) {
  const dataRange = piSheet.getDataRange();
  const values = dataRange.getValues();
  
  if (values.length < 4) {
    return [];
  }
  
  // Headers are in row 4 (index 3)
  const headers = values[3];
  
  // Find column indices
  const colMap = {};
  headers.forEach((header, index) => {
    colMap[header] = index;
  });
  
  const epics = [];
  
  // Process data rows (starting from row 5, index 4)
  for (let i = 4; i < values.length; i++) {
    const row = values[i];
    
    // Get value stream - check both columns
    const rowValueStream = row[colMap['Value Stream']] || row[colMap['Analyzed Value Stream']] || '';
    const issueType = row[colMap['Issue Type']] || '';
    
    // Skip if not matching value stream or not an Epic
    if (issueType !== 'Epic') continue;
    if (!rowValueStream.toString().toUpperCase().includes(valueStream.toUpperCase())) continue;
    
    const epic = {
      key: row[colMap['Key']] || '',
      summary: row[colMap['Summary']] || '',
      status: row[colMap['Status']] || '',
      valueStream: rowValueStream,
      scrumTeam: row[colMap['Scrum Team']] || '',
      allocation: row[colMap['Allocation']] || '',
      storyPoints: parseFloat(row[colMap['Story Points']]) || 0,
      storyPointEstimate: parseFloat(row[colMap['Story Point Estimate']]) || 0,
      featurePoints: parseFloat(row[colMap['Feature Points']]) || 0,
      portfolioInitiative: row[colMap['Portfolio Initiative']] || '',
      programInitiative: row[colMap['Program Initiative']] || '',
      piCommitment: row[colMap['PI Commitment']] || ''
    };
    
    epics.push(epic);
  }
  
  return epics;
}

/**
 * Get unique value streams from PI sheet
 * @param {Sheet} piSheet - The PI data sheet
 * @returns {Array} Array of unique value stream names
 */
function getUniqueValueStreamsFromPISheet(piSheet) {
  const dataRange = piSheet.getDataRange();
  const values = dataRange.getValues();
  
  if (values.length < 4) {
    return [];
  }
  
  const headers = values[3];
  const vsCol = headers.indexOf('Value Stream');
  const analyzedVsCol = headers.indexOf('Analyzed Value Stream');
  
  const valueStreams = new Set();
  
  for (let i = 4; i < values.length; i++) {
    const vs = values[i][vsCol] || values[i][analyzedVsCol] || '';
    if (vs && vs.toString().trim()) {
      valueStreams.add(vs.toString().trim());
    }
  }
  
  return Array.from(valueStreams).sort();
}

// ===== CHART FUNCTIONS =====

/**
 * Create a pie chart for initiative distribution
 * @param {Sheet} sheet - Target sheet
 * @param {number} dataStartRow - First row of data (not headers)
 * @param {number} dataRowCount - Number of data rows
 * @param {string} title - Chart title
 * @param {number} chartColumn - Column to position chart at
 * @param {number} chartRow - Row to position chart at
 */
function createInitiativePieChart(sheet, dataStartRow, dataRowCount, title, chartColumn, chartRow) {
  try {
    // Limit to top 10 for readability
    const rowCount = Math.min(dataRowCount, 10);
    
    const labelRange = sheet.getRange(dataStartRow, 1, rowCount, 1);  // Initiative names
    const valueRange = sheet.getRange(dataStartRow, 3, rowCount, 1);  // Points column
    
    const chart = sheet.newChart()
      .setChartType(Charts.ChartType.PIE)
      .addRange(labelRange)
      .addRange(valueRange)
      .setPosition(chartRow, chartColumn, 0, 0)
      .setOption('title', title)
      .setOption('width', INITIATIVE_ANALYSIS_CONFIG.chartWidth)
      .setOption('height', INITIATIVE_ANALYSIS_CONFIG.chartHeight)
      .setOption('pieSliceText', 'percentage')
      .setOption('legend', { position: 'right', textStyle: { fontSize: 9 } })
      .setOption('titleTextStyle', { fontSize: 12, bold: true })
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

// ===== INTEGRATION WITH MAIN ANALYSIS =====

/**
 * Add Initiative Analysis generation to the main analysis flow
 * Call this after the main PI analysis is complete
 * @param {number|string} piNumber - The PI number
 * @param {Array} valueStreams - Array of value stream names that were analyzed
 * @param {Spreadsheet} spreadsheet - The spreadsheet object
 */
function addInitiativeAnalysisToFlow(piNumber, valueStreams, spreadsheet) {
  console.log('Adding Initiative Analysis tabs to analysis flow...');
  
  valueStreams.forEach(vs => {
    generateInitiativeAnalysisForValueStream(piNumber, vs, spreadsheet);
  });
  
  console.log('Initiative Analysis tabs complete');
}

/**
 * Write Value Stream Capacity Distribution section
 * Aggregates capacity data across all teams in the value stream
 * @param {Sheet} sheet - The sheet to write to
 * @param {number} startRow - Starting row
 * @param {string} valueStream - The value stream name
 * @returns {number} Next available row
 */
function writeValueStreamCapacityDistribution(sheet, startRow, valueStream) {
  const colors = INITIATIVE_ANALYSIS_CONFIG.colors;
  const fontFamily = INITIATIVE_ANALYSIS_CONFIG.fontFamily;
  let currentRow = startRow;
  
  // Get all teams for this value stream
  const vsConfig = typeof VALUE_STREAM_CONFIG !== 'undefined' ? VALUE_STREAM_CONFIG[valueStream] : null;
  if (!vsConfig || !vsConfig.scrumTeams) {
    console.log(`No team config found for value stream ${valueStream}`);
    return currentRow;
  }
  
  const spreadsheet = sheet.getParent();
  
  // Aggregate capacity data across all teams
  const aggregatedCapacity = {
    productFeature: 0,
    productCompliance: 0,
    techPlatform: 0,
    quality: 0,
    klo: 0,
    unplannedWork: 0
  };
  
  let teamsWithData = 0;
  
  vsConfig.scrumTeams.forEach(team => {
    try {
      if (typeof getCapacityDataForTeamConsolidated === 'function') {
        const capacityData = getCapacityDataForTeamConsolidated(spreadsheet, team);
        if (capacityData && capacityData.allocations) {
          const allocs = capacityData.allocations;
          aggregatedCapacity.productFeature += allocs.productFeature || 0;
          aggregatedCapacity.productCompliance += allocs.productCompliance || 0;
          aggregatedCapacity.techPlatform += allocs.techPlatform || 0;
          aggregatedCapacity.quality += allocs.quality || 0;
          aggregatedCapacity.klo += allocs.klo || 0;
          aggregatedCapacity.unplannedWork += allocs.unplannedWork || 0;
          teamsWithData++;
        }
      }
    } catch (e) {
      console.log(`Error getting capacity for team ${team}: ${e.message}`);
    }
  });
  
  if (teamsWithData === 0) {
    console.log(`No capacity data found for any team in ${valueStream}`);
    return currentRow;
  }
  
  console.log(`Aggregated capacity data from ${teamsWithData} teams in ${valueStream}`);
  
  // Section header
  sheet.getRange(currentRow, 1).setValue('Value Stream Capacity Distribution');
  sheet.getRange(currentRow, 1, 1, 5)
    .setFontSize(14)
    .setFontWeight('bold')
    .setBackground(colors.purpleLight)
    .setFontFamily(fontFamily);
  sheet.getRange(currentRow, 1, 1, 5).merge();
  sheet.setRowHeight(currentRow, 28);
  currentRow++;
  
  sheet.getRange(currentRow, 1).setValue(`Capacity allocation across ${teamsWithData} teams`);
  sheet.getRange(currentRow, 1)
    .setFontStyle('italic')
    .setFontColor('#666666')
    .setFontSize(10)
    .setFontFamily(fontFamily);
  currentRow += 2;
  
  // Headers
  const headers = ['Allocation Type', 'Planned Capacity', '% of Total'];
  sheet.getRange(currentRow, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(currentRow, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground(colors.headerPrimary)
    .setFontColor(colors.white)
    .setFontSize(10)
    .setFontFamily(fontFamily)
    .setHorizontalAlignment('center');
  currentRow++;
  
  // Calculate totals
  const productCapacity = aggregatedCapacity.productFeature + aggregatedCapacity.productCompliance;
  const totalCapacity = productCapacity + aggregatedCapacity.techPlatform + 
                       aggregatedCapacity.quality + aggregatedCapacity.klo + 
                       aggregatedCapacity.unplannedWork;
  
  // Helper to calculate percentage
  const calcPercent = (value) => {
    if (totalCapacity === 0) return '0%';
    return Math.round((value / totalCapacity) * 100) + '%';
  };
  
  // Data rows
  const dataRows = [
    ['Product (Feature + Compliance)', Math.round(productCapacity), calcPercent(productCapacity)],
    ['Tech / Platform', Math.round(aggregatedCapacity.techPlatform), calcPercent(aggregatedCapacity.techPlatform)],
    ['Quality', Math.round(aggregatedCapacity.quality), calcPercent(aggregatedCapacity.quality)],
    ['KLO (Keep Lights On)', Math.round(aggregatedCapacity.klo), calcPercent(aggregatedCapacity.klo)],
    ['Unplanned Work', Math.round(aggregatedCapacity.unplannedWork), calcPercent(aggregatedCapacity.unplannedWork)]
  ];
  
  const dataStartRow = currentRow;
  
  dataRows.forEach((row, index) => {
    sheet.getRange(currentRow, 1, 1, row.length).setValues([row]);
    sheet.getRange(currentRow, 2, 1, 2).setHorizontalAlignment('center');
    
    if (index % 2 === 1) {
      sheet.getRange(currentRow, 1, 1, headers.length).setBackground('#f5f5f5');
    }
    currentRow++;
  });
  
  // Total row
  sheet.getRange(currentRow, 1, 1, 3).setValues([['TOTAL', Math.round(totalCapacity), '100%']]);
  sheet.getRange(currentRow, 1, 1, 3)
    .setFontWeight('bold')
    .setBackground(colors.goldAccent)
    .setFontSize(10)
    .setFontFamily(fontFamily);
  sheet.getRange(currentRow, 2, 1, 2).setHorizontalAlignment('center');
  currentRow++;
  
  // Create pie chart
  try {
    const chartColors = ['#1B365D', '#6B3FA0', '#4285F4', '#FFC72C', '#9AA0A6'];
    
    const labelRange = sheet.getRange(dataStartRow, 1, dataRows.length, 1);
    const valueRange = sheet.getRange(dataStartRow, 2, dataRows.length, 1);
    
    const chart = sheet.newChart()
      .setChartType(Charts.ChartType.PIE)
      .addRange(labelRange)
      .addRange(valueRange)
      .setPosition(dataStartRow - 2, 6, 0, 0)
      .setOption('title', `${valueStream} Capacity Distribution`)
      .setOption('width', 400)
      .setOption('height', 250)
      .setOption('pieSliceText', 'percentage')
      .setOption('legend', { position: 'right', textStyle: { fontSize: 10 } })
      .setOption('titleTextStyle', { fontSize: 12, bold: true })
      .setOption('colors', chartColors)
      .build();
    
    sheet.insertChart(chart);
  } catch (chartError) {
    console.log(`Could not create value stream pie chart: ${chartError.message}`);
  }
  
  return currentRow + 2;
}

/**
 * Test function to generate Initiative Analysis for a specific value stream
 */
function testInitiativeAnalysis() {
  const piNumber = 14;
  const valueStream = 'MMPM';
  
  console.log(`Testing Initiative Analysis for ${valueStream} in PI ${piNumber}`);
  
  const success = generateInitiativeAnalysisForValueStream(piNumber, valueStream);
  
  if (success) {
    console.log('Test successful!');
  } else {
    console.log('Test failed - check logs for errors');
  }
}