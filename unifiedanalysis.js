// ============================================================================
// UNIFIED ANALYSIS BACKEND
// ============================================================================
// This replaces the separate runAnalysisWithFilters + runAnalysisWithOptions
// functions with a single entry point from the new AnalysisDialog.
//
// WHAT CHANGED:
// - menuPrimaryUpdate: No longer shows a separate ui.prompt for PI number.
//   It now opens the unified AnalysisDialog directly.
// - menuAnalyzeWithSummaries / menuAnalyzeWithoutSummaries / menuAnalyzeWithSummaryChoice:
//   All three now open the same unified dialog with the summary option pre-selected.
// - menuAnalyzePICustom: Uses the new AnalysisDialog.html (already did, but now the
//   HTML is updated to include multi-select checkboxes and summary options).
// - runUnifiedAnalysis: Single backend function that handles PI selection, multi-select
//   value streams, summary options, AND report destination routing.
// - installRefreshTrigger: Stores analysis parameters in Document Properties so a
//   generated report spreadsheet can refresh itself.
//
// HOW TO APPLY:
// 1. Replace the AnalysisDialog HTML file with the new AnalysisDialog.html
// 2. Add this file (or merge these functions into menu.gs)
// 3. Remove / comment out the old functions:
//    - showValueStreamSelectionDialog(summaryOption)
//    - runAnalysisWithOptions(piNumber, selectedValueStreams, summaryOption)
//    - The two-popup version of menuPrimaryUpdate (lines ~104-339 in old menu.gs)
//    - runAnalysisWithFilters(params) -- replaced by runUnifiedAnalysis
// ============================================================================


/**
 * -- PRIMARY MENU ENTRY POINTS --
 * All of these now open the same unified dialog.
 * The only difference is which summary option is pre-selected.
 */

// "[TARGET] Analyze PI (with Destinations)..." -- opens full dialog
function menuAnalyzePICustom() {
  const html = HtmlService.createHtmlOutputFromFile('AnalysisDialog')
    .setWidth(520)
    .setHeight(720);
  SpreadsheetApp.getUi().showModalDialog(html, 'Analyze Program Increment');
}

// "Update Value Stream OR Scrum Team..." -- same unified dialog
// NOTE: If you still want the scrum-team-level update option, keep the old
// menuPrimaryUpdate and add a "Switch to Team Update" link in the dialog.
// For now this opens the unified dialog for value-stream-level updates.
function menuPrimaryUpdate() {
  const html = HtmlService.createHtmlOutputFromFile('AnalysisDialog')
    .setWidth(520)
    .setHeight(720);
  SpreadsheetApp.getUi().showModalDialog(html, 'Update PI Data');
}

// "Full Update (With Summaries)" -- same dialog, summary=with pre-selected
function menuAnalyzeWithSummaries() {
  const html = HtmlService.createHtmlOutputFromFile('AnalysisDialog')
    .setWidth(520)
    .setHeight(720);
  SpreadsheetApp.getUi().showModalDialog(html, 'Full Update (With Summaries)');
}

// "Fast Update (Skip Summaries)" -- same dialog, summary=without pre-selected
function menuAnalyzeWithoutSummaries() {
  const html = HtmlService.createHtmlOutputFromFile('AnalysisDialog')
    .setWidth(520)
    .setHeight(720);
  SpreadsheetApp.getUi().showModalDialog(html, 'Fast Update (Skip Summaries)');
}

// "Update (Ask About Summaries)" -- same dialog
function menuAnalyzeWithSummaryChoice() {
  const html = HtmlService.createHtmlOutputFromFile('AnalysisDialog')
    .setWidth(520)
    .setHeight(720);
  SpreadsheetApp.getUi().showModalDialog(html, 'Update PI Data');
}


// ============================================================================
// UNIFIED BACKEND FUNCTION
// ============================================================================

/**
 * Single backend function called from the unified AnalysisDialog.
 * Handles:
 *   - PI number validation
 *   - Multi-select value streams
 *   - Summary option (with / without)
 *   - Report destination (existing / new / update)
 *   - Saving refresh parameters for generated reports
 * 
 * @param {Object} params
 * @param {string} params.piNumber          - PI number (e.g. "15")
 * @param {string[]} params.valueStreams     - Array of selected value stream names
 * @param {string} params.summaryOption      - "with" or "without"
 * @param {string} params.reportDestination  - "existing", "new", or "update"
 * @param {string} [params.existingReportUrl] - Google Sheets URL (for "update")
 * @returns {string} Success message
 */
function runUnifiedAnalysis(params) {
  const { piNumber, valueStreams, summaryOption, reportDestination, existingReportUrl } = params;
  
  console.log('runUnifiedAnalysis called:', JSON.stringify(params));
  
  // -- Validation --
  if (!piNumber) throw new Error('PI number is required.');
  if (!valueStreams || valueStreams.length === 0) throw new Error('At least one value stream is required.');
  
  const updateSummaries = (summaryOption === 'with');
  
  let targetSpreadsheet = null;
  let reportUrl = '';
  let reportId = '';
  let reportName = '';
  
  try {
    // ================================================================
    // STEP 1: Resolve target spreadsheet based on destination
    // ================================================================
    
    switch (reportDestination) {
      
      case 'new': {
        showProgress('Creating new report spreadsheet...');
        
        const newName = 'PI ' + piNumber + ' - ' + valueStreams.join(', ') + ' Analysis';
        const newSpreadsheet = SpreadsheetApp.create(newName);
        
        targetSpreadsheet = newSpreadsheet;
        reportUrl = newSpreadsheet.getUrl();
        reportId = newSpreadsheet.getId();
        reportName = newName;
        
        // Move to same Drive folder as source spreadsheet
        try {
          const sourceFile = DriveApp.getFileById(
            SpreadsheetApp.getActiveSpreadsheet().getId()
          );
          const parents = sourceFile.getParents();
          if (parents.hasNext()) {
            const folder = parents.next();
            const newFile = DriveApp.getFileById(newSpreadsheet.getId());
            folder.addFile(newFile);
            DriveApp.getRootFolder().removeFile(newFile);
          }
        } catch (moveError) {
          console.warn('Could not move to source folder:', moveError);
        }
        
        // NOTE: Capacity Planning and Team Registry sheets are NOT copied to new reports
        // The new report references data from the PI sheet only
        
        // Remove default Sheet1
        try {
          const defaultSheet = targetSpreadsheet.getSheetByName('Sheet1');
          if (defaultSheet && targetSpreadsheet.getSheets().length > 1) {
            targetSpreadsheet.deleteSheet(defaultSheet);
          }
        } catch (e) { /* ignore */ }
        
        // -- Install refresh parameters --
        installRefreshParams_(targetSpreadsheet, params);
        
        console.log('Created new spreadsheet: ' + reportName + ' (' + reportId + ')');
        break;
      }
      
      case 'update': {
        if (!existingReportUrl) {
          throw new Error('Report URL is required for update mode.');
        }
        
        showProgress('Opening existing report spreadsheet...');
        
        var idMatch = existingReportUrl.match(/\/spreadsheets\/d\/([a-zA-Z0-9_-]+)/);
        if (!idMatch) throw new Error('Invalid Google Sheets URL format.');
        
        reportId = idMatch[1];
        
        try {
          targetSpreadsheet = SpreadsheetApp.openById(reportId);
        } catch (openError) {
          throw new Error(
            'Could not open the spreadsheet. Check the URL and your edit access.'
          );
        }
        
        reportUrl = existingReportUrl;
        reportName = targetSpreadsheet.getName();
        
        // -- Update refresh parameters --
        installRefreshParams_(targetSpreadsheet, params);
        
        console.log('Opened existing spreadsheet: ' + reportName + ' (' + reportId + ')');
        break;
      }
      
      case 'existing':
      default: {
        targetSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
        reportUrl = targetSpreadsheet.getUrl();
        reportId = targetSpreadsheet.getId();
        reportName = targetSpreadsheet.getName();
        
        console.log('Using active spreadsheet: ' + reportName);
        break;
      }
    }
    
    // ================================================================
    // STEP 2: Run the analysis
    // ================================================================
    
    showProgress('Running PI ' + piNumber + ' analysis for ' + valueStreams.join(', ') + '...');
    
    analyzeSelectedValueStreams(piNumber, valueStreams, {
      targetSpreadsheet: targetSpreadsheet,
      updateSummaries: false,  // We'll handle summaries ourselves below
      suppressCompletionAlert: true  // We'll show our own completion message
    });
    
    // ================================================================
    // STEP 2.5: Generate supporting artifacts if summaries requested
    // ================================================================
    
    let summaryResults = {
      teamSummaries: 0,
      initiativeAnalysis: 0,
      dansReport: false
    };
    
    if (updateSummaries) {
      const programIncrement = 'PI ' + piNumber;
      
      // Read the PI data sheet to get all issues
      showProgress('Reading PI data for summary generation...');
      const piSheet = targetSpreadsheet.getSheetByName(programIncrement);
      
      if (piSheet) {
        const dataRange = piSheet.getDataRange();
        const values = dataRange.getValues();
        
        if (values.length > 4) {
          const headers = values[3];
          const allIssues = parsePISheetData(values, headers);
          
          // Build set of valid teams for the selected value streams from config
          const validTeamsForSelectedVS = new Set();
          valueStreams.forEach(function(vs) {
            const vsConfig = VALUE_STREAM_CONFIG[vs];
            if (vsConfig && vsConfig.scrumTeams) {
              vsConfig.scrumTeams.forEach(function(team) {
                validTeamsForSelectedVS.add(team.toUpperCase());
              });
            }
          });
          
          console.log('Valid teams for selected value streams: ' + Array.from(validTeamsForSelectedVS).join(', '));
          
          // Filter issues: must match value stream AND team must be valid for that value stream
          const filteredIssues = allIssues.filter(function(issue) {
            const issueVS = (issue.analyzedValueStream || issue.valueStream || '').toUpperCase();
            const issueTeam = (issue.scrumTeam || '').toUpperCase();
            
            // Check value stream matches
            const vsMatches = valueStreams.some(function(vs) {
              return issueVS === vs.toUpperCase();
            });
            
            // Check team is valid for selected value streams (if team exists)
            const teamValid = !issue.scrumTeam || validTeamsForSelectedVS.has(issueTeam);
            
            return vsMatches && teamValid;
          });
          
          console.log('Total issues in PI sheet: ' + allIssues.length);
          console.log('Filtered issues for selected value streams: ' + filteredIssues.length);
          
          // Get unique scrum teams from the FILTERED issues only
          const teamsToGenerate = new Set();
          filteredIssues.forEach(function(issue) {
            if (issue.scrumTeam && validTeamsForSelectedVS.has(issue.scrumTeam.toUpperCase())) {
              teamsToGenerate.add(issue.scrumTeam);
            }
          });
          
          // Generate team summaries - SORT ALPHABETICALLY
          var teamsArray = Array.from(teamsToGenerate).sort(function(a, b) {
            return a.toLowerCase().localeCompare(b.toLowerCase());
          });
          console.log('Generating summaries for ' + teamsArray.length + ' teams (alphabetical): ' + teamsArray.join(', '));
          
          // Keep reference to source spreadsheet for capacity data
          const sourceSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
          
          teamsArray.forEach(function(team) {
            showProgress('Generating summary for ' + team + '...');
            try {
              var teamIssues = filteredIssues.filter(function(issue) { 
                return issue.scrumTeam === team; 
              });
              
              if (teamIssues.length > 0) {
                createScrumTeamSummary(teamIssues, programIncrement, team, targetSpreadsheet, sourceSpreadsheet);
                summaryResults.teamSummaries++;
                console.log('Created summary for ' + team);
              }
            } catch (teamError) {
              console.error('Error creating summary for ' + team + ':', teamError);
            }
          });
          
          // Generate Initiative Analysis tabs for each value stream
          valueStreams.forEach(function(vs) {
            showProgress('Generating Initiative Analysis for ' + vs + '...');
            try {
              if (typeof generateInitiativeAnalysisForValueStream === 'function') {
                var success = generateInitiativeAnalysisForValueStream(piNumber, vs, targetSpreadsheet);
                if (success) {
                  summaryResults.initiativeAnalysis++;
                  console.log('Created Initiative Analysis for ' + vs);
                }
              } else {
                console.log('generateInitiativeAnalysisForValueStream not available - skipping');
              }
            } catch (iaError) {
              console.error('Error creating Initiative Analysis for ' + vs + ':', iaError);
            }
          });
          
          // Generate Dan's Report ONLY if EMA Clinical is selected
          var hasEMAClinical = valueStreams.some(function(vs) { 
            return vs.toUpperCase() === 'EMA CLINICAL'; 
          });
          
          if (hasEMAClinical) {
            showProgress('Generating Dan\'s Report (Clinical Capacity)...');
            try {
              // Use the non-interactive version - pass SOURCE spreadsheet for capacity data
              if (typeof generateDansReportForSpreadsheet === 'function') {
                var success = generateDansReportForSpreadsheet(targetSpreadsheet, piNumber, sourceSpreadsheet);
                if (success) {
                  summaryResults.dansReport = true;
                  console.log('Created Dan\'s Report');
                } else {
                  console.log('Dan\'s Report generation returned false');
                }
              } else {
                console.log('generateDansReportForSpreadsheet not available - skipping');
              }
            } catch (dansError) {
              console.error('Error creating Dan\'s Report:', dansError);
            }
          }
        }
      }
    }
    
    // ================================================================
    // STEP 3: Log the report
    // ================================================================
    
    try {
      var epicCount = countEpicsInSpreadsheet_(targetSpreadsheet, piNumber);
      
      logReport({
        piNumber: piNumber,
        valueStream: valueStreams.join(', '),
        reportName: 'PI ' + piNumber + ' Analysis - ' + valueStreams.join(', '),
        spreadsheetUrl: reportUrl,
        spreadsheetId: reportId,
        epicCount: epicCount,
        status: 'Success'
      });
    } catch (logError) {
      console.error('Error logging report (non-critical):', logError);
    }
    
    // ================================================================
    // STEP 4: Return success message
    // ================================================================
    
    closeProgress();
    
    // Build detailed message
    var message = '';
    
    if (reportDestination === 'new') {
      message = 'Report created: "' + reportName + '"';
    } else if (reportDestination === 'update') {
      message = 'Report updated: "' + reportName + '"';
    } else {
      message = 'Analysis complete for PI ' + piNumber;
    }
    
    if (updateSummaries) {
      message += '\n\nGenerated:';
      message += '\n• Team Summaries: ' + summaryResults.teamSummaries;
      message += '\n• Initiative Analysis: ' + summaryResults.initiativeAnalysis;
      if (summaryResults.dansReport) {
        message += '\n• Dan\'s Report: Yes';
      }
    } else {
      message += ' (data only)';
    }
    
    // Return structured result for the dialog to parse
    if (reportDestination === 'new' || reportDestination === 'update') {
      // Return JSON string that the dialog can parse
      return JSON.stringify({
        success: true,
        message: message,
        reportUrl: reportUrl,
        reportName: reportName,
        summaryResults: summaryResults
      });
    } else {
      return message;
    }
    
  } catch (error) {
    console.error('Error in runUnifiedAnalysis:', error);
    
    // Log the failure
    try {
      logReport({
        piNumber: piNumber,
        valueStream: valueStreams.join(', '),
        reportName: 'PI ' + piNumber + ' Analysis - FAILED',
        spreadsheetUrl: reportUrl,
        spreadsheetId: reportId,
        epicCount: 0,
        status: 'Failed: ' + error.message
      });
    } catch (logError) {
      console.error('Error logging failure:', logError);
    }
    
    throw error;
  }
}


// ============================================================================
// REFRESH CAPABILITY FOR GENERATED REPORTS
// ============================================================================

/**
 * Stores analysis parameters in the target spreadsheet's Document Properties
 * so the report can be refreshed later without returning to the source sheet.
 * 
 * @param {Spreadsheet} targetSpreadsheet - The report spreadsheet
 * @param {Object} params - The original analysis parameters
 * @private
 */
function installRefreshParams_(targetSpreadsheet, params) {
  try {
    var props = PropertiesService.getDocumentProperties();
    
    // Store the source spreadsheet ID so we know where the data came from
    var sourceId = SpreadsheetApp.getActiveSpreadsheet().getId();
    
    var refreshConfig = {
      piNumber: params.piNumber,
      valueStreams: params.valueStreams,
      summaryOption: params.summaryOption,
      sourceSpreadsheetId: sourceId,
      lastRefreshed: new Date().toISOString()
    };
    
    props.setProperty('REFRESH_CONFIG', JSON.stringify(refreshConfig));
    
    console.log('Refresh parameters installed:', JSON.stringify(refreshConfig));
  } catch (e) {
    console.warn('Could not install refresh parameters:', e);
  }
}

/**
 * Called from the "Refresh" button on a generated report.
 * Reads stored parameters and re-runs the analysis.
 * 
 * This function is designed to be called from the report spreadsheet itself.
 * It uses Document Properties to retrieve the original analysis parameters.
 */
function refreshReportData() {
  var ui = SpreadsheetApp.getUi();
  
  try {
    var props = PropertiesService.getDocumentProperties();
    var configStr = props.getProperty('REFRESH_CONFIG');
    
    if (!configStr) {
      ui.alert(
        'No Refresh Configuration',
        'This spreadsheet does not have refresh parameters stored.\n\n' +
        'To set up refresh, run the analysis from the source spreadsheet ' +
        'using "Create New Report" or "Update Existing Report" destination.',
        ui.ButtonSet.OK
      );
      return;
    }
    
    var config = JSON.parse(configStr);
    
    // Confirm with user
    var response = ui.alert(
      'Refresh Report Data',
      'This will refresh the report with the following parameters:\n\n' +
      '• PI: ' + config.piNumber + '\n' +
      '• Value Streams: ' + config.valueStreams.join(', ') + '\n' +
      '• Summary: ' + (config.summaryOption === 'with' ? 'Full (with summaries)' : 'Fast (skip summaries)') + '\n' +
      '• Last refreshed: ' + (config.lastRefreshed ? new Date(config.lastRefreshed).toLocaleString() : 'Never') + '\n\n' +
      'Continue?',
      ui.ButtonSet.YES_NO
    );
    
    if (response !== ui.Button.YES) return;
    
    showProgress('Refreshing PI ' + config.piNumber + ' data...');
    
    var updateSummaries = (config.summaryOption === 'with');
    var targetSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    
    // Run the analysis targeting this spreadsheet
    analyzeSelectedValueStreams(config.piNumber, config.valueStreams, {
      targetSpreadsheet: targetSpreadsheet,
      updateSummaries: updateSummaries,
      suppressCompletionAlert: false
    });
    
    // Update last refreshed timestamp
    config.lastRefreshed = new Date().toISOString();
    props.setProperty('REFRESH_CONFIG', JSON.stringify(config));
    
    closeProgress();
    
    ui.alert(
      'Refresh Complete',
      'PI ' + config.piNumber + ' data has been refreshed for ' + config.valueStreams.join(', ') + '.',
      ui.ButtonSet.OK
    );
    
  } catch (error) {
    closeProgress();
    console.error('Error refreshing report:', error);
    ui.alert('Refresh Error', 'Failed to refresh: ' + error.message, ui.ButtonSet.OK);
  }
}

/**
 * Shows the stored refresh configuration for the current report.
 */
function showRefreshConfig() {
  var ui = SpreadsheetApp.getUi();
  
  try {
    var props = PropertiesService.getDocumentProperties();
    var configStr = props.getProperty('REFRESH_CONFIG');
    
    if (!configStr) {
      ui.alert('No refresh configuration found for this spreadsheet.');
      return;
    }
    
    var config = JSON.parse(configStr);
    
    ui.alert(
      'Refresh Configuration',
      'PI: ' + config.piNumber + '\n' +
      'Value Streams: ' + config.valueStreams.join(', ') + '\n' +
      'Summary: ' + config.summaryOption + '\n' +
      'Source Spreadsheet: ' + config.sourceSpreadsheetId + '\n' +
      'Last Refreshed: ' + (config.lastRefreshed || 'Never'),
      ui.ButtonSet.OK
    );
    
  } catch (e) {
    ui.alert('Error reading config: ' + e.message);
  }
}


// ============================================================================
// UPDATED onOpen -- ADDS REFRESH BUTTON IF ON A GENERATED REPORT
// ============================================================================

/**
 * Call this from your existing onOpen() to add a Refresh menu item
 * when the spreadsheet has refresh parameters stored.
 * 
 * Usage: Add this to your existing onOpen():
 *   addRefreshMenuIfNeeded();
 */
function addRefreshMenuIfNeeded() {
  try {
    var props = PropertiesService.getDocumentProperties();
    var configStr = props.getProperty('REFRESH_CONFIG');
    
    if (configStr) {
      var config = JSON.parse(configStr);
      var ui = SpreadsheetApp.getUi();
      
      ui.createMenu('Refresh Report')
        .addItem('Refresh Data (PI ' + config.piNumber + ')', 'refreshReportData')
        .addItem('View Refresh Config', 'showRefreshConfig')
        .addToUi();
    }
  } catch (e) {
    // Silently skip if properties aren't available
    console.log('Could not check for refresh config:', e);
  }
}