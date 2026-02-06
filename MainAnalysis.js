/**
 * Main Analysis Functions - Core functionality for PI analysis and summary generation
 * This file contains the main functions for analyzing JIRA data and generating reports
 */

// ===== MAIN ANALYSIS FUNCTIONS =====

/**
 * Generate summary for a specific value stream
 * @param {string} piNumber - The PI number
 * @param {string} valueStream - The value stream name
 */
function generateSummaryForValueStream(piNumber, valueStream) {
  const ui = SpreadsheetApp.getUi();
  const programIncrement = `PI ${piNumber}`;
  
  try {
    showProgress(`Generating summary for ${valueStream} in ${programIncrement}...`);
    
    // Check if PI data sheet exists
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const piSheetName = `PI ${piNumber}`;
    const piSheet = spreadsheet.getSheetByName(piSheetName);
    
    if (!piSheet) {
      closeProgress();
      ui.alert(`No data found for ${programIncrement}. Please run the analysis first.`);
      return;
    }
    
    // Read data from PI sheet
    showProgress('Reading PI data...');
    const dataRange = piSheet.getDataRange();
    const values = dataRange.getValues();
    const headers = values[3]; // Headers on row 4
    
    const allIssues = parsePISheetData(values, headers);
    
    // Filter for the selected value stream
    const vsIssues = allIssues.filter(issue => 
      issue.analyzedValueStream === valueStream || issue.valueStream === valueStream
    );
    
    if (vsIssues.length === 0) {
      closeProgress();
      ui.alert(`No data found for ${valueStream} in ${programIncrement}.`);
      return;
    }
    
    // Create summary sheet
    showProgress('Creating summary sheet...');
    const summarySheetName = `${programIncrement} - ${valueStream} Summary`;
    createValueStreamSummary(summarySheetName, vsIssues, programIncrement, valueStream);
    
    closeProgress();
    
    ui.alert(
      'Summary Complete',
      `Successfully generated summary for ${valueStream} in ${programIncrement}.\n\n` +
      `Sheet: "${summarySheetName}"`,
      ui.ButtonSet.OK
    );
    
  } catch (error) {
    console.error('Error generating summary:', error);
    closeProgress();
    ui.alert('Error', 'Summary generation failed: ' + error.toString(), ui.ButtonSet.OK);
  }
}


// ===== DATA PROCESSING FUNCTIONS =====

/**
 * Process epic data from JIRA response
 * @param {Object} epic - JIRA epic object
 * @param {string} analyzedValueStream - The value stream being analyzed
 * @return {Object} Processed epic data
 */
function processEpicData(epic, analyzedValueStream) {
  const fields = epic.fields;
  
  // Extract momentum label if present
  let momentum = '';
  if (fields.labels && Array.isArray(fields.labels)) {
    const momentumLabel = fields.labels.find(label => 
      label.toLowerCase().startsWith('momentum')
    );
    if (momentumLabel) {
      momentum = momentumLabel;
    }
  }
    
  // Calculate LOE (sum of story points from children will be added later)
  let loeEstimate = 0;
  
  return {
    key: epic.key,
    issueType: 'Epic',
    summary: fields.summary,
    status: fields.status?.name || '',
    valueStream: fields.customfield_10046 || '',
    scrumTeam: fields.customfield_10040 || '',
    allocation: fields.customfield_10043 || '',
    storyPoints: 0, // Epics don't have story points
    storyPointEstimate: parseFloat(fields.customfield_10016) || 0,
    epicLink: '',
    parentKey: '',
    featurePoints: parseFloat(fields.customfield_10252) || 0,
    loeEstimate: loeEstimate,
    programIncrement: fields.customfield_10113 || '',
    piCommitment: fields.customfield_10063 || '',
    components: (fields.components || []).map(c => c.name).join(', '),
    costOfDelay: fields.customfield_10065 || '',
    costOfDelay: parseFloat(fields.customfield_10065) || 0,
    dependsOnTeam: fields.customfield_10120 || '', 
    momentum: momentum
  };
}

/**
 * Process child issue data from JIRA response
 * @param {Object} child - JIRA child issue object
 * @param {string} epicKey - Parent epic key
 * @param {string} analyzedValueStream - The value stream being analyzed
 * @return {Object} Processed child data
 */
function processChildData(child, epicKey, analyzedValueStream) {
  const fields = child.fields;
  
  // Extract momentum label if present (though less common on child issues)
    let momentum = '';
    if (fields.labels && Array.isArray(fields.labels)) {
      const momentumLabel = fields.labels.find(label => 
        label.toLowerCase().startsWith('momentum')
      );
      if (momentumLabel) {
        momentum = momentumLabel;
      }
    }
  
  return {
    key: child.key,
    issueType: fields.issuetype?.name || '',
    summary: fields.summary,
    status: fields.status?.name || '',
    valueStream: fields.customfield_10046 || '',
    scrumTeam: fields.customfield_10040 || '',
    allocation: fields.customfield_10043 || '',
    storyPoints: parseFloat(fields.customfield_10037) || 0,
    storyPointEstimate: 0, // Children don't have estimates
    epicLink: fields.customfield_10014 || epicKey,
    parentKey: fields.parent?.key || '',
    featurePoints: 0, // Children don't have feature points
    loeEstimate: 0, // Children don't have LOE
    programIncrement: fields.customfield_10113 || '',
    piCommitment: fields.customfield_10063 || '',
        components: (fields.components || []).map(c => c.name).join(', '),
    costOfDelay: 0,  // Children don't have CoD,
    dependsOnTeam: fields.customfield_10120 || '',
    momentum: momentum  
  };
}

// ===== SHEET CREATION FUNCTIONS =====

/**
 * Create or update PI analysis sheet
 * @param {string} sheetName - Name of the sheet
 * @param {Array} issues - Array of issue objects
 * @param {Array} analyzedValueStreams - Array of analyzed value streams
 */
function createPIAnalysisSheet(sheetName, issues, analyzedValueStreams) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName(sheetName);
  
  if (sheet) {
    sheet.clear();
  } else {
    sheet = spreadsheet.insertSheet(sheetName);
  }
  
  const headers = [
    'Key',
    'Issue Type',
    'Summary',
    'Status',
    'Value Stream',
    'Scrum Team',
    'Allocation',
    'Story Points',
    'Story Point Estimate',
    'Epic Link',
    'Parent',
    'Feature Points',
    'LOE Estimate',
    'Program Increment',
    'PI Commitment',
    'Components',
    'Cost of Delay', 
    'Momentum',
    'Depends on Valuestream',
    'Depends on Team'
  ];
  
  sheet.getRange(1, 1).setValue(`PI Analysis - ${sheetName}`);
  sheet.getRange(1, 1).setFontSize(16).setFontWeight('bold');
  
  sheet.getRange(2, 1).setValue('Last Updated:');
  sheet.getRange(2, 2).setValue(new Date().toLocaleString());
  sheet.getRange(2, 1, 1, 2).setFontWeight('bold');
  
  sheet.getRange(3, 1).setValue('Analyzed Value Streams:');
  sheet.getRange(3, 2).setValue(analyzedValueStreams.join(', '));
  sheet.getRange(3, 1, 1, 2).setFontWeight('bold');
  
  sheet.getRange(4, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(4, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#4285f4')
    .setFontColor('white');
  
  const epicLOE = calculateEpicLOE(issues);
  
  const data = issues.map(issue => {
    if (issue.issueType === 'Epic' && epicLOE[issue.key]) {
      issue.loeEstimate = epicLOE[issue.key];
    }
    
    return [
      issue.key,
      issue.issueType,
      issue.summary,
      issue.status,
      issue.valueStream,
      issue.scrumTeam,
      issue.allocation,
      issue.storyPoints,
      issue.storyPointEstimate,
      issue.epicLink,
      issue.parentKey,
      issue.featurePoints,
      issue.loeEstimate,
      issue.programIncrement,
      issue.piCommitment,
      issue.components,
      issue.costOfDelay || 0,
      issue.momentum || '',
      issue.dependsOnValuestream || '',
      issue.dependsOnTeam || ''
    ];
  });
  
  if (data.length > 0) {
    sheet.getRange(5, 1, data.length, headers.length).setValues(data);
    
    const keys = data.map(row => row[0]);
    applyJiraHyperlinks(sheet, 5, 1, keys);
  }
  
  sheet.setFrozenRows(4);
  sheet.setFrozenColumns(1);
  sheet.autoResizeColumns(1, headers.length);
  
  sheet.setColumnWidth(3, 400);
  sheet.getRange(5, 3, data.length, 1).setWrap(true);
}

/**
 * Create value stream summary sheet
 * @param {string} sheetName - Name of the summary sheet
 * @param {Array} issues - Array of issue objects
 * @param {string} programIncrement - PI name
 * @param {string} valueStream - Value stream name
 */
function createValueStreamSummary(sheetName, issues, programIncrement, valueStream) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName(sheetName);
  
  if (sheet) {
    sheet.clear();
  } else {
    sheet = spreadsheet.insertSheet(sheetName);
  }
  
  // Separate epics and stories
  const epics = issues.filter(i => i.issueType === 'Epic');
  const stories = issues.filter(i => i.issueType !== 'Epic');
  
  // Set up summary header
  sheet.getRange(1, 1).setValue(`${programIncrement} - ${valueStream} Summary`);
  sheet.getRange(1, 1).setFontSize(16).setFontWeight('bold');
  
  sheet.getRange(2, 1).setValue('Last Updated:');
  sheet.getRange(2, 2).setValue(new Date().toLocaleString());
  sheet.getRange(2, 1, 1, 2).setFontWeight('bold');
  
  let currentRow = 4;
  
  // Summary metrics
  sheet.getRange(currentRow, 1).setValue('Summary Metrics');
  sheet.getRange(currentRow, 1).setFontSize(14).setFontWeight('bold').setBackground('#e8f0fe');
  currentRow += 2;
  
  const metrics = [
    ['Total Epics', epics.length],
    ['Total Stories/Tasks', stories.length],
    ['Total Story Points', stories.reduce((sum, s) => sum + (s.storyPoints || 0), 0)],
    ['Total Story Point Estimates', epics.reduce((sum, e) => sum + (e.storyPointEstimate || 0), 0)],
    ['Total Feature Points (x10)', epics.reduce((sum, e) => sum + ((e.featurePoints || 0) * 10), 0)]
  ];
  
  metrics.forEach(metric => {
    sheet.getRange(currentRow, 1).setValue(metric[0]);
    sheet.getRange(currentRow, 2).setValue(metric[1]);
    currentRow++;
  });
  
  currentRow += 2;
  
  // Epics section
  sheet.getRange(currentRow, 1).setValue('Epics');
  sheet.getRange(currentRow, 1).setFontSize(14).setFontWeight('bold').setBackground('#e8f0fe');
  currentRow += 2;
  
  if (epics.length > 0) {
    const epicHeaders = [
      'Key', 'Summary', 'Status', 'Scrum Team', 'Allocation',
      'Story Point Estimate', 'LOE Estimate', 'Feature Points (x10)'
    ];
    
    sheet.getRange(currentRow, 1, 1, epicHeaders.length).setValues([epicHeaders]);
    sheet.getRange(currentRow, 1, 1, epicHeaders.length)
      .setFontWeight('bold')
      .setBackground('#4285f4')
      .setFontColor('white');
    currentRow++;
    
    const epicData = epics.map(epic => [
      epic.key,
      epic.summary.substring(0, 60) + (epic.summary.length > 60 ? '...' : ''),
      epic.status,
      epic.scrumTeam || 'Unassigned',
      epic.allocation || '',
      epic.storyPointEstimate || 0,
      epic.loeEstimate || 0,
      (epic.featurePoints || 0) * 10
    ]);
    
    sheet.getRange(currentRow, 1, epicData.length, epicHeaders.length).setValues(epicData);
    
    // Apply hyperlinks
    const epicKeys = epics.map(e => e.key);
    applyJiraHyperlinks(sheet, currentRow, 1, epicKeys);
    
    currentRow += epicData.length;
  }
  
  // Format the sheet
  sheet.setFrozenRows(4);
  sheet.autoResizeColumns(1, 8);
  sheet.setColumnWidth(2, 400); // Summary column
}

// ===== HELPER FUNCTIONS =====

/**
 * Calculate LOE (Level of Effort) for epics
 * @param {Array} issues - All issues
 * @return {Object} Map of epic keys to LOE values
 */
function calculateEpicLOE(issues) {
  const epicLOE = {};
  
  // Initialize all epics with 0 LOE
  issues.filter(i => i.issueType === 'Epic').forEach(epic => {
    epicLOE[epic.key] = 0;
  });
  
  // Sum story points from children
  issues.filter(i => i.issueType !== 'Epic').forEach(child => {
    const epicKey = child.epicLink || child.parentKey;
    if (epicKey && epicLOE.hasOwnProperty(epicKey)) {
      epicLOE[epicKey] += child.storyPoints || 0;
    }
  });
  
  return epicLOE;
}