// MenuIntegration.gs
// PLANNING Tool - Menu Integration for PI Planning with Scrum Team Focus
// ===== MENU SETUP - UPDATE FOCUSED =====
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSheet();
  const sheetName = sheet.getName();
  
  // Main PLANNING Tool menu with Update as primary action
  const planningMenu = ui.createMenu('ðŸ“‹ PLANNING Tool');
  
  // PRIMARY UPDATE ACTIONS
  planningMenu.addItem('ðŸ”„ Update Value Stream OR Scrum Team...', 'menuPrimaryUpdate');
  planningMenu.addSeparator();
  
  // Analysis and Reporting
  planningMenu.addSubMenu(ui.createMenu('ðŸ“Š Analysis & Reports')
      .addItem('ðŸŽ¯ Analyze PI (with Destinations)...', 'menuAnalyzePICustom')
      .addSeparator()
      .addItem('ðŸ“ˆ Full Update (With Summaries)', 'menuAnalyzeWithSummaries')
      .addItem('âš¡ Fast Update (Skip Summaries)', 'menuAnalyzeWithoutSummaries')
      .addItem('â“ Update (Ask About Summaries)', 'menuAnalyzeWithSummaryChoice')
    .addItem('Generate Dans Report', 'generateDansReport')
    .addSeparator()
    .addItem('ðŸ“ˆ View PI Planning Dashboard', 'menuShowPIPlanningDashboard'));
  
  planningMenu.addSeparator();
  
  // Advanced Operations
  planningMenu.addSubMenu(ui.createMenu('ðŸš€ Advanced Operations')
    .addItem('Batch Update Multiple Teams...', 'menuBatchUpdateTeams')
    .addItem('Update All Teams in Value Stream...', 'menuUpdateAllTeamsInValueStream')
    .addItem('Analyze All Value Streams...', 'menuAnalyzePIAllValueStreams')
    .addSeparator()
    .addItem('Refresh All Value Streams from JIRA...', 'menuRefreshAllValueStreams'));
  
  planningMenu.addSeparator();
  
  // Utilities submenu
  planningMenu.addSubMenu(ui.createMenu('ðŸ”§ Utilities')
    .addItem('Test JIRA Connection', 'menuTestJiraConnection')
    .addItem('Configure JIRA Credentials', 'showCredentialSetupDialog')
    .addSeparator()
    .addItem('Refresh All Formulas', 'menuRefreshFormulas')
    .addItem('Clear Cache', 'menuClearCache')
    .addItem('Clean Current Sheet Data', 'cleanCurrentSheetData')
    .addSeparator()
    .addItem('Show All Scrum Teams', 'menuShowAllScrumTeams')
    .addItem('Show All Value Streams', 'menuShowAllValueStreams')
    .addItem('ðŸ” Diagnostic Check', 'diagnosticCheckDansReportData')
    .addSeparator()
    .addItem('ðŸ“‹ View Report Log', 'viewReportLog')
    .addItem('ðŸ“‚ Open Report from Log...', 'openReportFromLog')
    .addSeparator()
    .addItem('Setup Instructions', 'menuSetup'));
    
  planningMenu.addToUi();
  
  // Add the filter menu
  const filterMenu = updateFilterMenu();
  filterMenu.addToUi();
  
  // Add refresh menu if this is a generated report
  addRefreshMenuIfNeeded();
}

function updateFilterMenu() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSheet();
  const sheetName = sheet.getName();
  
  const filterMenu = ui.createMenu('ðŸ” Quick Filters');
  
  // If on a summary sheet, add direct filter options
  if (sheetName.includes('Summary')) {
    filterMenu.addItem('âš ï¸ LOE Exceeding Estimate', 'instructLOEFilter');
    filterMenu.addItem('ðŸ”„ Allocation Mismatches', 'instructMismatchFilter');
    filterMenu.addItem('ðŸ“Š All Epics', 'instructAllEpicsFilter');
    
    // Add Stories filter for scrum team summaries
    if (!sheetName.includes(' - ') || sheetName.split(' - ').length > 2) {
      filterMenu.addItem('ðŸ“ Stories/Tasks by Status', 'instructStoriesFilter');
    }
    
    filterMenu.addSeparator();
    filterMenu.addItem('âŒ Clear All Filters', 'clearSheetFilters');
  } else {
    // If not on a summary sheet, add navigation options
    filterMenu.addItem('âš ï¸ Apply LOE Filter...', 'selectSummaryForLOEFilter');
    filterMenu.addItem('ðŸ”„ Apply Allocation Mismatch Filter...', 'selectSummaryForMismatchFilter');
    filterMenu.addItem('ðŸ“Š Apply All Epics Filter...', 'selectSummaryForAllEpicsFilter');
    filterMenu.addSeparator();
    filterMenu.addItem('ðŸ“‹ Go to Summary Sheet...', 'navigateToSummarySheet');
  }
  
  filterMenu.addSeparator();
  filterMenu.addItem('â“ Filter Instructions', 'showFilterInstructions');
  
  return filterMenu;
}

// =============================================================================
// PRIMARY MENU ENTRY POINTS - These are now defined in UnifiedAnalysis.gs
// The functions below (menuPrimaryUpdate, menuAnalyzePICustom, menuAnalyzeWithSummaries,
// menuAnalyzeWithoutSummaries, menuAnalyzeWithSummaryChoice) all open the unified
// AnalysisDialog.html which provides:
//   - PI number selection
//   - Multi-select value streams (with RCM Genie included)
//   - Summary options (with/without)
//   - Report destination (existing/new/update)
// =============================================================================

// NOTE: The following functions are defined in UnifiedAnalysis.gs:
// - menuPrimaryUpdate()
// - menuAnalyzePICustom()
// - menuAnalyzeWithSummaries()
// - menuAnalyzeWithoutSummaries()
// - menuAnalyzeWithSummaryChoice()
// - runUnifiedAnalysis(params)
// - addRefreshMenuIfNeeded()
// - refreshReportData()
// - showRefreshConfig()

/**
 * Batch update multiple teams
 */
function menuBatchUpdateTeams() {
  const ui = SpreadsheetApp.getUi();
  
  // First, get the PI number
  const piResponse = ui.prompt(
    'Batch Update Teams - Step 1',
    'Enter PI number (e.g., 11, 12, 13):',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (piResponse.getSelectedButton() !== ui.Button.OK) return;
  
  const piNumber = piResponse.getResponseText().trim();
  if (!piNumber || !/^\d+$/.test(piNumber)) {
    ui.alert('Invalid PI format. Please use a number like "11" or "12"');
    return;
  }
  
  // Get available scrum teams
  showProgress('Reading scrum teams...');
  const scrumTeams = getScrumTeamsFromPISheet(piNumber);
  closeProgress();
  
  if (scrumTeams.length === 0) {
    ui.alert('No scrum teams found in the PI data. Please run an analysis first.');
    return;
  }
  
  // Show team selection dialog
  const html = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: Arial, sans-serif; padding: 20px; }
      h3 { margin-top: 0; color: #333; }
      .info { background: #e8f0fe; padding: 10px; border-radius: 4px; margin-bottom: 15px; font-size: 14px; }
      .checkbox-container { max-height: 300px; overflow-y: auto; border: 1px solid #ddd; padding: 10px; margin: 10px 0; }
      .checkbox-item { margin: 5px 0; }
      .button-container { margin-top: 15px; text-align: right; }
      button { padding: 8px 15px; margin-left: 10px; font-size: 14px; cursor: pointer; }
      .update-btn { background-color: #34a853; color: white; border: none; }
      .update-btn:hover { background-color: #2d8e47; }
      .update-btn:disabled { background-color: #ccc; cursor: not-allowed; }
      .cancel-btn { background-color: #f1f3f4; color: #5f6368; border: 1px solid #dadce0; }
      .select-all { margin: 10px 0; font-size: 14px; }
      .processing { display: none; color: #34a853; margin-top: 10px; text-align: center; }
      .warning { background: #fff3cd; padding: 10px; border-radius: 4px; margin-top: 10px; font-size: 12px; color: #856404; }
    </style>
    
    <h3>Batch Update Scrum Teams for PI ${piNumber}</h3>
    <div class="info">
      Select multiple teams to update their data from JIRA
    </div>
    
    <div class="select-all">
      <label>
        <input type="checkbox" id="selectAll" onchange="toggleAll(this)"> 
        Select All Teams
      </label>
    </div>
    <div class="checkbox-container">
      ${scrumTeams.map((team, index) => `
        <div class="checkbox-item">
          <label>
            <input type="checkbox" class="team-checkbox" value="${team.name}"> 
            ${team.name} (${team.epicCount} epics, ${team.storyCount} stories)${team.name === 'Unassigned' ? ' - <em>Uses existing data only</em>' : ''}
          </label>
        </div>
      `).join('')}
    </div>
    
    <div class="warning">
      âš ï¸ Note: Updating multiple teams may take several minutes depending on the amount of data.
    </div>
    
    <div class="processing">Updating teams... This may take several minutes.</div>
    <div class="button-container">
      <button class="cancel-btn" onclick="google.script.host.close()">Cancel</button>
      <button class="update-btn" onclick="updateTeams()">Update Selected Teams</button>
    </div>
    
    <script>
      function toggleAll(checkbox) {
        const checkboxes = document.querySelectorAll('.team-checkbox');
        checkboxes.forEach(cb => cb.checked = checkbox.checked);
      }
      
      function updateTeams() {
        const selected = [];
        document.querySelectorAll('.team-checkbox:checked').forEach(cb => {
          selected.push(cb.value);
        });
        
        if (selected.length === 0) {
          alert('Please select at least one team');
          return;
        }
        
        if (selected.length > 5) {
          if (!confirm('You selected ' + selected.length + ' teams. This may take several minutes. Continue?')) {
            return;
          }
        }
        
        document.querySelector('.update-btn').disabled = true;
        document.querySelector('.processing').style.display = 'block';
        
        google.script.run
          .withSuccessHandler(() => {
            google.script.host.close();
          })
          .withFailureHandler(err => {
            alert('Error: ' + err);
            document.querySelector('.update-btn').disabled = false;
            document.querySelector('.processing').style.display = 'none';
          })
          .batchUpdateScrumTeams('${piNumber}', selected);
      }
    </script>
  `)
  .setWidth(500)
  .setHeight(550);
  
  ui.showModalDialog(html, 'Batch Update Teams');
}

/**
 * Update all teams in a value stream
 */
function menuUpdateAllTeamsInValueStream() {
  const ui = SpreadsheetApp.getUi();
  
  // First, get the PI number
  const piResponse = ui.prompt(
    'Update All Teams in Value Stream - Step 1',
    'Enter PI number (e.g., 11, 12, 13):',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (piResponse.getSelectedButton() !== ui.Button.OK) return;
  
  const piNumber = piResponse.getResponseText().trim();
  if (!piNumber || !/^\d+$/.test(piNumber)) {
    ui.alert('Invalid PI format. Please use a number like "13" or "14"');
    return;
  }
  
  const valueStreams = getAvailableValueStreams();
  
  // Show value stream selection
  const html = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: Arial, sans-serif; padding: 20px; }
      h3 { margin-top: 0; color: #333; }
      .info { background: #e8f0fe; padding: 10px; border-radius: 4px; margin-bottom: 15px; font-size: 14px; }
      select { width: 100%; padding: 8px; margin: 10px 0; font-size: 14px; }
      .button-container { margin-top: 20px; text-align: right; }
      button { padding: 8px 15px; margin-left: 10px; font-size: 14px; cursor: pointer; }
      .update-btn { background-color: #34a853; color: white; border: none; }
      .update-btn:disabled { background-color: #ccc; cursor: not-allowed; }
      .cancel-btn { background-color: #f1f3f4; color: #5f6368; border: 1px solid #dadce0; }
      .processing { display: none; color: #34a853; margin-top: 10px; text-align: center; }
    </style>
    
    <h3>Update All Teams in Value Stream</h3>
    <div class="info">
      <strong>PI ${piNumber}</strong> - This will update all teams in the selected value stream
    </div>
    
    <label>Select Value Stream:</label>
    <select id="valueStream" size="6">
      ${valueStreams.map(vs => `<option value="${vs}">${vs}</option>`).join('')}
    </select>
    
    <div class="processing">Updating all teams... This may take several minutes.</div>
    <div class="button-container">
      <button class="cancel-btn" onclick="google.script.host.close()">Cancel</button>
      <button class="update-btn" onclick="updateValueStream()">Update Value Stream</button>
    </div>
    
    <script>
      function updateValueStream() {
        const selected = document.getElementById('valueStream').value;
        
        if (!selected) {
          alert('Please select a value stream');
          return;
        }
        
        document.querySelector('.update-btn').disabled = true;
        document.querySelector('.processing').style.display = 'block';
        
        google.script.run
          .withSuccessHandler(() => {
            google.script.host.close();
          })
          .withFailureHandler(err => {
            alert('Error: ' + err);
            document.querySelector('.update-btn').disabled = false;
            document.querySelector('.processing').style.display = 'none';
          })
          .startAnalysisWrapper('${piNumber}', [selected]);
      }
    </script>
  `)
  .setWidth(450)
  .setHeight(400);
  
  ui.showModalDialog(html, 'Update Value Stream');
}

// ===== SCRUM TEAM FOCUSED FUNCTIONS =====

/**
 * Menu function to generate a scrum team summary
 */
function menuGenerateScrumTeamSummary() {
  const ui = SpreadsheetApp.getUi();
  
  // First, get the PI number
  const piResponse = ui.prompt(
    'Generate Scrum Team Summary - Step 1',
    'Enter PI number (e.g., 11, 12, 13):',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (piResponse.getSelectedButton() !== ui.Button.OK) return;
  
  const piNumber = piResponse.getResponseText().trim();
  if (!piNumber || !/^\d+$/.test(piNumber)) {
    ui.alert('Invalid PI format. Please use a number like "13" or "14"');
    return;
  }
  
  // Get available scrum teams
  showProgress('Reading scrum teams...');
  const scrumTeams = getScrumTeamsFromPISheet(piNumber);
  closeProgress();
  
  if (scrumTeams.length === 0) {
    ui.alert('No scrum teams found in the PI data. Please run an analysis first.');
    return;
  }
  
  // Show team selection dialog
  const html = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: Arial, sans-serif; padding: 20px; }
      h3 { margin-top: 0; color: #333; }
      .checkbox-container { max-height: 300px; overflow-y: auto; border: 1px solid #ddd; padding: 10px; margin: 10px 0; }
      .checkbox-item { margin: 5px 0; }
      .button-container { margin-top: 15px; text-align: right; }
      button { padding: 8px 15px; margin-left: 10px; font-size: 14px; cursor: pointer; }
      .generate-btn { background-color: #4285f4; color: white; border: none; }
      .generate-btn:hover { background-color: #357ae8; }
      .generate-btn:disabled { background-color: #ccc; cursor: not-allowed; }
      .cancel-btn { background-color: #f1f3f4; color: #5f6368; border: 1px solid #dadce0; }
      .select-all { margin: 10px 0; font-size: 14px; }
      .processing { display: none; color: #4285f4; margin-top: 10px; text-align: center; }
    </style>
    
    <h3>Generate Scrum Team Summary for PI ${piNumber}</h3>
    <div class="select-all">
      <label>
        <input type="checkbox" id="selectAll" onchange="toggleAll(this)"> 
        Select All Teams
      </label>
    </div>
    <div class="checkbox-container">
      ${scrumTeams.map((team, index) => `
        <div class="checkbox-item">
          <label>
            <input type="checkbox" class="team-checkbox" value="${team.name}"> 
            ${team.name} (${team.epicCount} epics, ${team.storyCount} stories)
          </label>
        </div>
      `).join('')}
    </div>
    
    <div class="processing">Generating summaries... This may take a moment.</div>
    <div class="button-container">
      <button class="cancel-btn" onclick="google.script.host.close()">Cancel</button>
      <button class="generate-btn" onclick="generateSummaries()">Generate Summaries</button>
    </div>
    
    <script>
      function toggleAll(checkbox) {
        const checkboxes = document.querySelectorAll('.team-checkbox');
        checkboxes.forEach(cb => cb.checked = checkbox.checked);
      }
      
      function generateSummaries() {
        const selected = [];
        document.querySelectorAll('.team-checkbox:checked').forEach(cb => {
          selected.push(cb.value);
        });
        
        if (selected.length === 0) {
          alert('Please select at least one scrum team');
          return;
        }
        
        document.querySelector('.generate-btn').disabled = true;
        document.querySelector('.processing').style.display = 'block';
        
        google.script.run
          .withSuccessHandler(() => {
            google.script.host.close();
          })
          .withFailureHandler(err => {
            alert('Error: ' + err);
            document.querySelector('.generate-btn').disabled = false;
            document.querySelector('.processing').style.display = 'none';
          })
          .generateScrumTeamSummaries('${piNumber}', selected);
      }
    </script>
  `)
  .setWidth(500)
  .setHeight(500);
  
  ui.showModalDialog(html, 'Select Scrum Teams');
}

/**
 * Refresh all value streams from JIRA
 */
function menuRefreshAllValueStreams() {
  menuAnalyzePIWithSelection();
}

/**
 * Show all available scrum teams
 */
function menuShowAllScrumTeams() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    showProgress('Discovering all scrum teams across all JIRA projects...');
    const scrumTeams = discoverScrumTeamsFromJira();
    closeProgress();
    
    // Get known teams from config
    const knownTeams = {};
    Object.keys(VALUE_STREAM_CONFIG).forEach(vs => {
      const teams = VALUE_STREAM_CONFIG[vs].scrumTeams || [];
      teams.forEach(team => {
        knownTeams[team] = vs;
      });
    });
    
    // Format the display
    let message = `Found ${scrumTeams.length} scrum teams across ALL projects in JIRA:\n\n`;
    
    scrumTeams.forEach(team => {
      if (knownTeams[team]) {
        message += `${team} (${knownTeams[team]})\n`;
      } else {
        message += `${team}\n`;
      }
    });
    
    message += '\n\nTeams marked with value streams are configured in the system.';
    
    ui.alert('Scrum Teams Found', message, ui.ButtonSet.OK);
    
  } catch (error) {
    closeProgress();
    ui.alert('Error', 'Failed to discover scrum teams: ' + error.toString(), ui.ButtonSet.OK);
  }
}

/**
 * Show all available value streams
 */
function menuShowAllValueStreams() {
  const ui = SpreadsheetApp.getUi();
  const valueStreams = getAvailableValueStreams();
  
  let message = `Configured Value Streams (${valueStreams.length}):\n\n`;
  
  valueStreams.forEach(vs => {
    const config = VALUE_STREAM_CONFIG[vs];
    message += `${vs}\n`;
    
    if (config && config.filter) {
      if (vs === 'AIMM') {
        message += `  - Searches for: ${config.filter.valueStreams.join(', ')}\n`;
        message += `  - With Scrum Team: ${config.filter.scrumTeam}\n`;
      }
    }
    
    if (config && config.scrumTeams && config.scrumTeams.length > 0) {
      message += `  - Known teams: ${config.scrumTeams.length}\n`;
    }
    
    message += '\n';
  });
  
  ui.alert('Value Streams', message, ui.ButtonSet.OK);
}

/**
 * Show PI Planning Dashboard
 */
function menuShowPIPlanningDashboard() {
  const ui = SpreadsheetApp.getUi();
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // Get current PI
  const piNumber = getCurrentPIFromSheets();
  if (!piNumber) {
    ui.alert('No PI data found. Please run an analysis first.');
    return;
  }
  
  // Create or update dashboard sheet
  showProgress('Creating PI Planning Dashboard...');
  createPIPlanningDashboard(piNumber);
  closeProgress();
  
  // Navigate to dashboard
  const dashboardSheet = spreadsheet.getSheetByName(`PI ${piNumber} Dashboard`);
  if (dashboardSheet) {
    spreadsheet.setActiveSheet(dashboardSheet);
    ui.alert(
      'Dashboard Created',
      `PI ${piNumber} Planning Dashboard has been created/updated.\n\n` +
      'The dashboard shows team-level and value stream metrics.',
      ui.ButtonSet.OK
    );
  }
}

// ===== LEGACY MENU FUNCTIONS (for backward compatibility) =====

/**
 * Legacy function - now redirects to unified dialog
 * This maintains backward compatibility if called directly
 */
function menuAnalyzePIWithSelection() {
  // Use the unified dialog instead
  menuAnalyzePICustom();
}

function menuGenerateSummaryCustom() {
  var ui = SpreadsheetApp.getUi();
  
  var piResponse = ui.prompt(
    'Generate Summary Report',
    'Enter PI number (e.g., 11, 12, 13):',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (piResponse.getSelectedButton() !== ui.Button.OK) return;
  
  var piNumber = piResponse.getResponseText().trim();
  if (!piNumber || !/^\d+$/.test(piNumber)) {
    ui.alert('Invalid PI format. Please use a number like "11" or "12"');
    return;
  }
  
  const availableValueStreams = getAvailableValueStreams();
  
  const html = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: Arial, sans-serif; padding: 20px; }
      select { width: 100%; padding: 8px; margin: 10px 0; font-size: 14px; }
      .button-container { margin-top: 20px; text-align: right; }
      button { padding: 8px 15px; margin-left: 10px; font-size: 14px; cursor: pointer; }
      .analyze-btn { background-color: #4285f4; color: white; border: none; }
      .analyze-btn:disabled { background-color: #ccc; cursor: not-allowed; }
      .cancel-btn { background-color: #f1f3f4; color: #5f6368; border: 1px solid #dadce0; }
      .processing { display: none; color: #4285f4; margin-top: 10px; text-align: center; }
      .note { font-size: 12px; color: #666; margin-top: 10px; }
    </style>
    
    <h3>Select Value Stream for PI ${piNumber} Summary</h3>
    <select id="valueStream" size="10">
      ${availableValueStreams.map(vs => `<option value="${vs}">${vs}</option>`).join('')}
    </select>
    
    <div class="note">
      Note: Summary will include data from ALL projects in JIRA
    </div>
    
    <div class="processing">Processing... This window will close automatically.</div>
    <div class="button-container">
      <button class="cancel-btn" onclick="google.script.host.close()">Cancel</button>
      <button class="analyze-btn" onclick="generate()">Generate Summary</button>
    </div>
    
    <script>
      function generate() {
        const select = document.getElementById('valueStream');
        const selected = select.value;
        
        if (!selected) {
          alert('Please select a value stream');
          return;
        }
        
        document.querySelector('.analyze-btn').disabled = true;
        document.querySelector('.processing').style.display = 'block';
        
        google.script.run
          .withSuccessHandler(() => {
            google.script.host.close();
          })
          .withFailureHandler(err => {
            alert('Error: ' + err);
            document.querySelector('.analyze-btn').disabled = false;
            document.querySelector('.processing').style.display = 'none';
          })
          .startSummaryWrapper('${piNumber}', selected);
      }
    </script>
  `)
  .setWidth(400)
  .setHeight(400);
  
  ui.showModalDialog(html, 'Select Value Stream');
}

// ===== ADVANCED MENU FUNCTIONS =====

function menuAnalyzePIAllValueStreams() {
  const ui = SpreadsheetApp.getUi();
  
  const response = ui.prompt(
    'Analyze PI - All Value Streams',
    'Enter PI number to analyze (e.g., 11, 12, 13):',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() === ui.Button.OK) {
    const piNumber = response.getResponseText().trim();
    
    if (piNumber && /^\d+$/.test(piNumber)) {
      try {
        showProgress('Discovering all value streams across all projects...');
        const allValueStreams = discoverValueStreamsFromJira();
        closeProgress();
        
        if (allValueStreams.length === 0) {
          ui.alert('No value streams found in JIRA.');
          return;
        }
        
        const confirmResponse = ui.alert(
          'Confirm Analysis',
          `Found ${allValueStreams.length} value streams across all projects:\n\n${allValueStreams.join(', ')}\n\nProceed with analysis?`,
          ui.ButtonSet.YES_NO
        );
        
        if (confirmResponse === ui.Button.YES) {
          analyzeSelectedValueStreams(piNumber, allValueStreams);
        }
        
      } catch (error) {
        console.error('Error:', error);
        ui.alert('âŒ Error: ' + error.message);
      }
    } else {
      ui.alert('Invalid PI format. Please use a number like "11" or "12"');
    }
  }
}

function menuGenerateSummaryFromCurrentData() {
  const ui = SpreadsheetApp.getUi();
  
  const piNumber = getCurrentPIFromSheets();
  if (!piNumber) {
    ui.alert('No PI data found. Please run an analysis first.');
    return;
  }
  
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const piSheet = spreadsheet.getSheetByName(`PI ${piNumber}`);
  
  if (!piSheet) {
    ui.alert('No PI data found. Please run an analysis first.');
    return;
  }
  
  const dataRange = piSheet.getDataRange();
  const values = dataRange.getValues();
  const headerRow = 3;
  
  if (values.length <= headerRow) {
    ui.alert('Invalid data format in PI sheet.');
    return;
  }
  
  const headers = values[headerRow];
  const analyzedVSIndex = headers.indexOf('Analyzed Value Stream');
  
  if (analyzedVSIndex === -1) {
    ui.alert('Cannot find value stream data in PI sheet.');
    return;
  }
  
  const valueStreams = new Set();
  for (let i = headerRow + 1; i < values.length; i++) {
    const vs = values[i][analyzedVSIndex];
    if (vs) valueStreams.add(vs);
  }
  
  const valueStreamArray = Array.from(valueStreams).sort();
  
  if (valueStreamArray.length === 0) {
    ui.alert('No value streams found in current data.');
    return;
  }
  
  const html = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: Arial, sans-serif; padding: 20px; }
      h3 { margin-top: 0; color: #333; }
      .info { background: #e8f0fe; padding: 10px; border-radius: 4px; margin-bottom: 15px; font-size: 14px; }
      select { width: 100%; padding: 8px; margin: 10px 0; font-size: 14px; }
      .button-container { margin-top: 20px; text-align: right; }
      button { padding: 8px 15px; margin-left: 10px; font-size: 14px; cursor: pointer; }
      .generate-btn { background-color: #4285f4; color: white; border: none; }
      .generate-btn:disabled { background-color: #ccc; cursor: not-allowed; }
      .cancel-btn { background-color: #f1f3f4; color: #5f6368; border: 1px solid #dadce0; }
      .processing { display: none; color: #4285f4; margin-top: 10px; text-align: center; }
    </style>
    
    <h3>Generate Summary from Current Data</h3>
    <div class="info">
      Current PI: <strong>PI ${piNumber}</strong><br>
      Found ${valueStreamArray.length} value stream(s) in data<br>
      <em>Data includes epics from all projects</em>
    </div>
    
    <label>Select Value Stream:</label>
    <select id="valueStream" size="8">
      ${valueStreamArray.map(vs => `<option value="${vs}">${vs}</option>`).join('')}
    </select>
    
    <div class="processing">Processing... This window will close automatically.</div>
    <div class="button-container">
      <button class="cancel-btn" onclick="google.script.host.close()">Cancel</button>
      <button class="generate-btn" onclick="generate()">Generate Summary</button>
    </div>
    
    <script>
      function generate() {
        const select = document.getElementById('valueStream');
        const selected = select.value;
        
        if (!selected) {
          alert('Please select a value stream');
          return;
        }
        
        document.querySelector('.generate-btn').disabled = true;
        document.querySelector('.processing').style.display = 'block';
        
        google.script.run
          .withSuccessHandler(() => {
            google.script.host.close();
          })
          .withFailureHandler(err => {
            alert('Error: ' + err);
            document.querySelector('.generate-btn').disabled = false;
            document.querySelector('.processing').style.display = 'none';
          })
          .startSummaryWrapper('${piNumber}', selected);
      }
    </script>
  `)
  .setWidth(400)
  .setHeight(420);
  
  ui.showModalDialog(html, 'Generate Summary');
}

function menuClearCache() {
  const ui = SpreadsheetApp.getUi();
  
  const response = ui.alert(
    'Clear Cache',
    'What would you like to clear?\n\n' +
    'â€¢ All Caches - Remove all cached data\n' +
    'â€¢ Specific PI - Clear cache for a specific PI\n' +
    'â€¢ Cancel - Keep existing cache',
    ui.ButtonSet.YES_NO_CANCEL
  );
  
  if (response === ui.Button.YES) {
    // Clear all
    try {
      CacheManager.clearAll();
      ui.alert('âœ… All caches cleared successfully!');
    } catch (error) {
      ui.alert('âŒ Error clearing cache: ' + error.message);
    }
  } else if (response === ui.Button.NO) {
    // Clear specific PI
    const piResponse = ui.prompt(
      'Clear PI Cache',
      'Enter PI number to clear (e.g., 11, 12):',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (piResponse.getSelectedButton() === ui.Button.OK) {
      const piNumber = piResponse.getResponseText().trim();
      if (piNumber && /^\d+$/.test(piNumber)) {
        try {
          CacheManager.clearPI(piNumber);
          ui.alert(`âœ… Cache cleared for PI ${piNumber}!`);
        } catch (error) {
          ui.alert('âŒ Error clearing cache: ' + error.message);
        }
      }
    }
  }
}

function menuSetup() {
  setup();
}

// ===== PLACEHOLDER FOR TEST JIRA CONNECTION =====
function menuTestJiraConnection() {
  const ui = SpreadsheetApp.getUi();
  ui.alert(
    'Test JIRA Connection',
    'This function needs to be implemented in the main JIRA integration file.',
    ui.ButtonSet.OK
  );
}

// ===== FILTER MENU FUNCTIONS =====

/**
 * Get filter views for current sheet (requires permissions)
 */
function getFilterViewsForCurrentSheet() {
  const spreadsheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
  const sheet = SpreadsheetApp.getActiveSheet();
  const currentSheetId = sheet.getSheetId();
  
  try {
    const response = Sheets.Spreadsheets.get(spreadsheetId, {
      fields: 'sheets(properties(sheetId,title),filterViews)'
    });
    
    let filterViews = [];
    response.sheets.forEach(sheetData => {
      if (sheetData.properties.sheetId === currentSheetId && sheetData.filterViews) {
        filterViews = sheetData.filterViews;
      }
    });
    
    return filterViews;
  } catch (error) {
    console.log('Could not load filter views:', error);
    return [];
  }
}

// ===== NEW FILTER NAVIGATION FUNCTIONS =====

function instructLOEFilter() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const sheetName = sheet.getName();
  
  if (!sheetName.includes('Summary')) {
    SpreadsheetApp.getUi().alert('This filter only works on summary sheets.');
    return;
  }
  
  // Show immediate processing feedback
  const processingToast = SpreadsheetApp.getActiveSpreadsheet().toast('â³ Applying LOE filter...', 'Processing', -1);
  
  try {
    // Clear any existing filters first
    const existingFilter = sheet.getFilter();
    if (existingFilter) {
      try {
        existingFilter.remove();
      } catch (e) {
        // Filter might already be removed, continue
      }
    }
    
    // Small delay to ensure filter is cleared
    Utilities.sleep(100);
    
    // Find the LOE Exceeding Estimate section
    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    
    let loeStartRow = -1;
    let headerRow = -1;
    let numColumns = 8; // Default columns for LOE section
    
    // Find the section - look for variations of the header text
    for (let i = 0; i < values.length; i++) {
      const cellValue = values[i][0];
      if (cellValue && (
          cellValue.toString().includes('Epics where LOE Estimate exceeds Story Point Estimate') ||
          cellValue.toString().includes('Epics where LOE exceeds') ||
          cellValue.toString().includes('LOE Exceeding') ||
          (cellValue.toString() === 'Epics' && i > 10) // For scrum team summaries
      )) {
        // Found the section header, look for the table headers
        for (let j = i + 1; j < values.length && j < i + 5; j++) {
          if (values[j][0] === 'Key') {
            headerRow = j + 1; // Convert to 1-based
            loeStartRow = j + 2; // Data starts after header
            
            // Count actual columns in this section
            numColumns = 0;
            for (let col = 0; col < values[j].length; col++) {
              if (values[j][col] && values[j][col] !== '') {
                numColumns++;
              } else {
                break;
              }
            }
            break;
          }
        }
        if (headerRow !== -1) break;
      }
    }
    
    if (headerRow === -1) {
      SpreadsheetApp.getActiveSpreadsheet().toast('', '', 1); // Clear toast
      SpreadsheetApp.getUi().alert('Could not find Epic/LOE section in this sheet.');
      return;
    }
    
    // Find the end of this section - look for empty row or next section
    let loeEndRow = loeStartRow;
    for (let i = loeStartRow - 1; i < values.length; i++) {
      // Check if we hit an empty row or a new section
      if (!values[i][0] || values[i][0] === '') {
        loeEndRow = i + 1; // Convert to 1-based
        break;
      }
      // Check for next section headers
      const cellText = values[i][0].toString();
      if (i > loeStartRow - 1 && 
          (cellText.includes('Allocation Mismatches') || 
           cellText.includes('All Epics') ||
           cellText.includes('Stories/Tasks') ||
           cellText.includes('Color Key'))) {
        loeEndRow = i + 1; // Convert to 1-based
        break;
      }
    }
    
    // If we didn't find an end, check up to 100 rows
    if (loeEndRow === loeStartRow) {
      loeEndRow = Math.min(loeStartRow + 100, sheet.getMaxRows());
    }
    
    console.log(`LOE Filter: Header at row ${headerRow}, Data from ${loeStartRow} to ${loeEndRow}, Columns: ${numColumns}`);
    
    // Apply filter to the LOE section
    const filterRange = sheet.getRange(headerRow, 1, loeEndRow - headerRow, numColumns);
    const filter = filterRange.createFilter();
    
    // Find the difference column (usually named "Difference")
    const headerValues = sheet.getRange(headerRow, 1, 1, numColumns).getValues()[0];
    let diffColumn = headerValues.indexOf('Difference') + 1; // Convert to 1-based
    if (diffColumn === 0) {
      // If not found, assume it's one of the last columns
      diffColumn = Math.max(7, numColumns - 1);
    }
    
    // Sort by difference column in descending order
    filter.sort(diffColumn, false);
    
    // Optionally, hide rows where difference <= 0
    if (diffColumn > 0) {
      const criteria = SpreadsheetApp.newFilterCriteria()
        .whenNumberGreaterThan(0)
        .build();
      filter.setColumnFilterCriteria(diffColumn, criteria);
    }
    
    // Clear the processing toast and show success
    SpreadsheetApp.getActiveSpreadsheet().toast('âœ… LOE filter applied successfully!', 'Complete', 3);
    
  } catch (error) {
    console.error('Error applying LOE filter:', error);
    SpreadsheetApp.getActiveSpreadsheet().toast('', '', 1); // Clear toast
    
    // Check if it's a "filter already exists" error
    if (error.toString().includes('already has a filter')) {
      SpreadsheetApp.getActiveSpreadsheet().toast('Filter is already active. Use Clear All Filters first.', 'Info', 3);
    } else {
      SpreadsheetApp.getUi().alert('Error applying filter: ' + error.toString());
    }
  }
}

/**
 * Apply Allocation Mismatches filter - updated to handle when section doesn't exist
 */
function instructMismatchFilter() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const sheetName = sheet.getName();
  
  if (!sheetName.includes('Summary')) {
    SpreadsheetApp.getUi().alert('This filter only works on summary sheets.');
    return;
  }
  
  // Show immediate processing feedback
  const processingToast = SpreadsheetApp.getActiveSpreadsheet().toast('â³ Applying Allocation Mismatch filter...', 'Processing', -1);
  
  try {
    // Clear any existing filters first
    const existingFilter = sheet.getFilter();
    if (existingFilter) {
      try {
        existingFilter.remove();
      } catch (e) {
        // Filter might already be removed, continue
      }
    }
    
    // Small delay to ensure filter is cleared
    Utilities.sleep(100);
    
    // Find the Allocation Mismatches section
    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    
    let mismatchStartRow = -1;
    let headerRow = -1;
    let numColumns = 5; // Base columns
    
    // Find the section
    for (let i = 0; i < values.length; i++) {
      const cellValue = values[i][0];
      if (cellValue && cellValue.toString().includes('Allocation Mismatches')) {
        // Found the section header, look for the table headers
        for (let j = i + 1; j < values.length && j < i + 5; j++) {
          if (values[j][0] === 'Key') {
            headerRow = j + 1; // Convert to 1-based
            mismatchStartRow = j + 2; // Data starts after header
            
            // Count actual columns (including Mismatch columns)
            numColumns = 0;
            for (let col = 0; col < values[j].length; col++) {
              if (values[j][col] && values[j][col] !== '') {
                numColumns++;
              }
            }
            break;
          }
        }
        break;
      }
    }
    
    if (headerRow === -1) {
      SpreadsheetApp.getActiveSpreadsheet().toast('', '', 1); // Clear toast
      SpreadsheetApp.getUi().alert(
        'No Allocation Mismatches Found',
        'This sheet does not have an Allocation Mismatches section.\n' +
        'This could mean there are no allocation conflicts in this data.',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }
    
    // Find the end of this section
    let mismatchEndRow = mismatchStartRow;
    for (let i = mismatchStartRow - 1; i < values.length; i++) {
      if (!values[i][0] || values[i][0] === '' || 
          (typeof values[i][0] === 'string' && 
           (values[i][0].includes('All Epics') || 
            values[i][0].includes('Stories/Tasks') ||
            values[i][0].includes('Epics') && i > mismatchStartRow))) {
        mismatchEndRow = i + 1;
        break;
      }
    }
    
    if (mismatchEndRow === mismatchStartRow) {
      mismatchEndRow = Math.min(mismatchStartRow + 50, sheet.getMaxRows());
    }
    
    console.log(`Mismatch Filter: Header at row ${headerRow}, Data from ${mismatchStartRow} to ${mismatchEndRow}, Columns: ${numColumns}`);
    
    // Apply filter to the Mismatches section
    const filterRange = sheet.getRange(headerRow, 1, mismatchEndRow - headerRow, numColumns);
    const filter = filterRange.createFilter();
    
    // Sort by Scrum Team (column 3) if it exists, otherwise by allocation
    const headerValues = sheet.getRange(headerRow, 1, 1, numColumns).getValues()[0];
    let sortColumn = headerValues.indexOf('Scrum Team') + 1;
    if (sortColumn === 0) {
      sortColumn = headerValues.indexOf('Allocation') + 1;
    }
    if (sortColumn === 0) {
      sortColumn = 3; // Default to column 3
    }
    
    filter.sort(sortColumn, true);
    
    // Clear the processing toast and show success
    SpreadsheetApp.getActiveSpreadsheet().toast('âœ… Allocation Mismatch filter applied successfully!', 'Complete', 3);
    
  } catch (error) {
    console.error('Error applying mismatch filter:', error);
    SpreadsheetApp.getActiveSpreadsheet().toast('', '', 1); // Clear toast
    
    // Check if it's a "filter already exists" error
    if (error.toString().includes('already has a filter')) {
      SpreadsheetApp.getActiveSpreadsheet().toast('Filter is already active. Use Clear All Filters first.', 'Info', 3);
    } else {
      SpreadsheetApp.getUi().alert('Error applying filter: ' + error.toString());
    }
  }
}

/**
 * Apply All Epics filter - updated for both summary types
 */
function instructAllEpicsFilter() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const sheetName = sheet.getName();
  
  if (!sheetName.includes('Summary')) {
    SpreadsheetApp.getUi().alert('This filter only works on summary sheets.');
    return;
  }
  
  // Show immediate processing feedback
  const processingToast = SpreadsheetApp.getActiveSpreadsheet().toast('â³ Applying All Epics filter...', 'Processing', -1);
  
  try {
    // Clear any existing filters first
    const existingFilter = sheet.getFilter();
    if (existingFilter) {
      try {
        existingFilter.remove();
      } catch (e) {
        // Filter might already be removed, continue
      }
    }
    
    // Small delay to ensure filter is cleared
    Utilities.sleep(100);
    
    // Find the All Epics section or Epics section
    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    
    let epicsStartRow = -1;
    let headerRow = -1;
    let numColumns = 8; // Default columns for epics
    
    // Find the section - look for the last occurrence of "All Epics" or just "Epics"
    for (let i = values.length - 1; i >= 0; i--) {
      const cellValue = values[i][0];
      if (cellValue && (cellValue.toString() === 'All Epics' || 
                       (cellValue.toString() === 'Epics' && i > 10))) {
        // Found the section header, look for the table headers
        for (let j = i + 1; j < values.length && j < i + 5; j++) {
          if (values[j][0] === 'Key') {
            headerRow = j + 1; // Convert to 1-based
            epicsStartRow = j + 2; // Data starts after header
            
            // Count actual columns
            numColumns = 0;
            for (let col = 0; col < values[j].length; col++) {
              if (values[j][col] && values[j][col] !== '') {
                numColumns++;
              } else {
                break;
              }
            }
            break;
          }
        }
        break;
      }
    }
    
    if (headerRow === -1) {
      SpreadsheetApp.getActiveSpreadsheet().toast('', '', 1); // Clear toast
      SpreadsheetApp.getUi().alert('Could not find Epics section in this sheet.');
      return;
    }
    
    // For epics section, go to the last row with data in column A or next section
    let epicsEndRow = epicsStartRow;
    for (let i = epicsStartRow - 1; i < values.length; i++) {
      if (!values[i][0] || values[i][0] === '' ||
          (typeof values[i][0] === 'string' && values[i][0].includes('Stories/Tasks'))) {
        epicsEndRow = i + 1;
        break;
      }
    }
    
    if (epicsEndRow === epicsStartRow) {
      epicsEndRow = sheet.getMaxRows() + 1;
    }
    
    console.log(`All Epics Filter: Header at row ${headerRow}, Data from ${epicsStartRow} to ${epicsEndRow}, Columns: ${numColumns}`);
    
    // Apply filter to the Epics section
    const filterRange = sheet.getRange(headerRow, 1, epicsEndRow - headerRow, numColumns);
    const filter = filterRange.createFilter();
    
    // Sort by Scrum Team or Allocation
    const headerValues = sheet.getRange(headerRow, 1, 1, numColumns).getValues()[0];
    let teamColumn = headerValues.indexOf('Scrum Team') + 1;
    let allocationColumn = headerValues.indexOf('Allocation') + 1;
    
    if (teamColumn > 0) {
      filter.sort(teamColumn, true);
    } else if (allocationColumn > 0) {
      filter.sort(allocationColumn, true);
    }
    
    // Clear the processing toast and show success
    SpreadsheetApp.getActiveSpreadsheet().toast('âœ… All Epics filter applied successfully!', 'Complete', 3);
    
  } catch (error) {
    console.error('Error applying all epics filter:', error);
    SpreadsheetApp.getActiveSpreadsheet().toast('', '', 1); // Clear toast
    
    // Check if it's a "filter already exists" error
    if (error.toString().includes('already has a filter')) {
      SpreadsheetApp.getActiveSpreadsheet().toast('Filter is already active. Use Clear All Filters first.', 'Info', 3);
    } else {
      SpreadsheetApp.getUi().alert('Error applying filter: ' + error.toString());
    }
  }
}

/**
 * Add a new filter for Stories/Tasks section (for scrum team summaries)
 */
function instructStoriesFilter() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const sheetName = sheet.getName();
  
  if (!sheetName.includes('Summary')) {
    SpreadsheetApp.getUi().alert('This filter only works on summary sheets.');
    return;
  }
  
  // Show immediate processing feedback
  const processingToast = SpreadsheetApp.getActiveSpreadsheet().toast('â³ Applying Stories filter...', 'Processing', -1);
  
  try {
    // Clear any existing filters first
    const existingFilter = sheet.getFilter();
    if (existingFilter) {
      try {
        existingFilter.remove();
      } catch (e) {
        // Filter might already be removed, continue
      }
    }
    
    // Small delay to ensure filter is cleared
    Utilities.sleep(100);
    
    // Find the Stories/Tasks section
    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    
    let storiesStartRow = -1;
    let headerRow = -1;
    let numColumns = 3; // Default columns for stories section
    
    // Find the section
    for (let i = 0; i < values.length; i++) {
      const cellValue = values[i][0];
      if (cellValue && cellValue.toString().includes('Stories/Tasks')) {
        // Found the section header, look for the table headers
        for (let j = i + 1; j < values.length && j < i + 5; j++) {
          if (values[j][0] === 'Status') {
            headerRow = j + 1; // Convert to 1-based
            storiesStartRow = j + 2; // Data starts after header
            
            // Count actual columns
            numColumns = 0;
            for (let col = 0; col < values[j].length; col++) {
              if (values[j][col] && values[j][col] !== '') {
                numColumns++;
              } else {
                break;
              }
            }
            break;
          }
        }
        break;
      }
    }
    
    if (headerRow === -1) {
      SpreadsheetApp.getActiveSpreadsheet().toast('', '', 1); // Clear toast
      SpreadsheetApp.getUi().alert('Could not find Stories/Tasks section in this sheet.');
      return;
    }
    
    // Find the end of this section
    let storiesEndRow = sheet.getMaxRows() + 1;
    for (let i = storiesStartRow - 1; i < values.length; i++) {
      if (!values[i][0] || values[i][0] === '') {
        storiesEndRow = i + 1;
        break;
      }
    }
    
    console.log(`Stories Filter: Header at row ${headerRow}, Data from ${storiesStartRow} to ${storiesEndRow}, Columns: ${numColumns}`);
    
    // Apply filter to the Stories section
    const filterRange = sheet.getRange(headerRow, 1, storiesEndRow - headerRow, numColumns);
    const filter = filterRange.createFilter();
    
    // Sort by Status (first column)
    filter.sort(1, true);
    
    // Clear the processing toast and show success
    SpreadsheetApp.getActiveSpreadsheet().toast('âœ… Stories filter applied successfully!', 'Complete', 3);
    
  } catch (error) {
    console.error('Error applying stories filter:', error);
    SpreadsheetApp.getActiveSpreadsheet().toast('', '', 1); // Clear toast
    SpreadsheetApp.getUi().alert('Error applying filter: ' + error.toString());
  }
}

function clearSheetFilters() {
  // Show immediate processing feedback
  const processingToast = SpreadsheetApp.getActiveSpreadsheet().toast('â³ Clearing filters...', 'Processing', -1);
  
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    const filter = sheet.getFilter();
    
    if (filter) {
      filter.remove();
      SpreadsheetApp.getActiveSpreadsheet().toast('âœ… All filters have been cleared!', 'Complete', 3);
    } else {
      SpreadsheetApp.getActiveSpreadsheet().toast('â„¹ï¸ No active filters to clear on this sheet.', 'Info', 3);
    }
  } catch (error) {
    SpreadsheetApp.getActiveSpreadsheet().toast('', '', 1); // Clear toast
    console.error('Error clearing filters:', error);
    SpreadsheetApp.getActiveSpreadsheet().toast('âš ï¸ Error clearing filters: ' + error.toString(), 'Error', 5);
  }
}

/**
 * Shows a dialog to select a summary sheet for LOE filter
 */
function selectSummaryForLOEFilter() {
  selectSummarySheet('LOE', 'instructLOEFilter');
}

/**
 * Shows a dialog to select a summary sheet for Mismatch filter
 */
function selectSummaryForMismatchFilter() {
  selectSummarySheet('Mismatch', 'instructMismatchFilter');
}

/**
 * Shows a dialog to select a summary sheet for All Epics filter
 */
function selectSummaryForAllEpicsFilter() {
  selectSummarySheet('AllEpics', 'instructAllEpicsFilter');
}

/**
 * Generic function to select a summary sheet and apply a filter
 */
function selectSummarySheet(filterType, filterFunction) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = spreadsheet.getSheets();
  const summarySheets = sheets.filter(sheet => sheet.getName().includes('Summary'));
  
  if (summarySheets.length === 0) {
    SpreadsheetApp.getUi().alert(
      'No Summary Sheets Found',
      'No summary sheets found. Please generate a summary report first.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }
  
  // If only one summary sheet, go directly to it
  if (summarySheets.length === 1) {
    spreadsheet.setActiveSheet(summarySheets[0]);
    // Call the appropriate filter function
    this[filterFunction]();
    return;
  }
  
  // Multiple summary sheets - show selection dialog
  const html = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: Arial, sans-serif; padding: 20px; }
      h3 { margin-top: 0; }
      select { width: 100%; padding: 8px; margin: 10px 0; font-size: 14px; }
      .button-container { margin-top: 20px; text-align: right; }
      button { padding: 8px 15px; margin-left: 10px; font-size: 14px; cursor: pointer; }
      .apply-btn { background-color: #4285f4; color: white; border: none; }
      .apply-btn:disabled { background-color: #ccc; cursor: not-allowed; }
      .cancel-btn { background-color: #f1f3f4; color: #5f6368; border: 1px solid #dadce0; }
    </style>
    
    <h3>Select Summary Sheet</h3>
    <p>Choose a summary sheet to apply the filter:</p>
    <select id="sheetName" size="8">
      ${summarySheets.map(sheet => `<option value="${sheet.getName()}">${sheet.getName()}</option>`).join('')}
    </select>
    
    <div class="button-container">
      <button class="cancel-btn" onclick="google.script.host.close()">Cancel</button>
      <button class="apply-btn" onclick="applyFilter()">Apply Filter</button>
    </div>
    
    <script>
      function applyFilter() {
        const select = document.getElementById('sheetName');
        const selected = select.value;
        
        if (!selected) {
          alert('Please select a sheet');
          return;
        }
        
        // Disable button and show processing
        document.querySelector('.apply-btn').disabled = true;
        document.querySelector('.apply-btn').textContent = 'Applying...';
        
        google.script.run
          .withSuccessHandler(() => {
            google.script.host.close();
          })
          .withFailureHandler(err => {
            alert('Error: ' + err);
            document.querySelector('.apply-btn').disabled = false;
            document.querySelector('.apply-btn').textContent = 'Apply Filter';
          })
          .applyFilterToSheet(selected, '${filterFunction}');
      }
    </script>
  `)
  .setWidth(400)
  .setHeight(350);
  
  const ui = SpreadsheetApp.getUi();
  ui.showModalDialog(html, 'Select Summary Sheet');
}

/**
 * Applies a filter to a specific sheet
 */
function applyFilterToSheet(sheetName, filterFunction) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(sheetName);
  
  if (!sheet) {
    throw new Error('Sheet not found: ' + sheetName);
  }
  
  // Activate the sheet
  spreadsheet.setActiveSheet(sheet);
  
  // Call the appropriate filter function
  this[filterFunction]();
}

/**
 * Navigate to a summary sheet
 */
function navigateToSummarySheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = spreadsheet.getSheets();
  const summarySheets = sheets.filter(sheet => sheet.getName().includes('Summary'));
  
  if (summarySheets.length === 0) {
    SpreadsheetApp.getUi().alert(
      'No Summary Sheets Found',
      'No summary sheets found. Please generate a summary report first.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }
  
  // If only one summary sheet, go directly to it
  if (summarySheets.length === 1) {
    spreadsheet.setActiveSheet(summarySheets[0]);
    SpreadsheetApp.getUi().alert(
      'Navigation Complete',
      `Navigated to: ${summarySheets[0].getName()}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }
  
  // Multiple summary sheets - show selection dialog
  const html = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: Arial, sans-serif; padding: 20px; }
      h3 { margin-top: 0; }
      select { width: 100%; padding: 8px; margin: 10px 0; font-size: 14px; }
      .button-container { margin-top: 20px; text-align: right; }
      button { padding: 8px 15px; margin-left: 10px; font-size: 14px; cursor: pointer; }
      .go-btn { background-color: #4285f4; color: white; border: none; }
      .cancel-btn { background-color: #f1f3f4; color: #5f6368; border: 1px solid #dadce0; }
    </style>
    
    <h3>Select Summary Sheet</h3>
    <select id="sheetName" size="8">
      ${summarySheets.map(sheet => `<option value="${sheet.getName()}">${sheet.getName()}</option>`).join('')}
    </select>
    
    <div class="button-container">
      <button class="cancel-btn" onclick="google.script.host.close()">Cancel</button>
      <button class="go-btn" onclick="goToSheet()">Go to Sheet</button>
    </div>
    
    <script>
      function goToSheet() {
        const select = document.getElementById('sheetName');
        const selected = select.value;
        
        if (!selected) {
          alert('Please select a sheet');
          return;
        }
        
        google.script.run
          .withSuccessHandler(() => {
            google.script.host.close();
          })
          .withFailureHandler(err => {
            alert('Error: ' + err);
          })
          .navigateToSheet(selected);
      }
    </script>
  `)
  .setWidth(400)
  .setHeight(350);
  
  const ui = SpreadsheetApp.getUi();
  ui.showModalDialog(html, 'Select Summary Sheet');
}

/**
 * Navigate to a specific sheet
 */
function navigateToSheet(sheetName) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(sheetName);
  
  if (!sheet) {
    throw new Error('Sheet not found: ' + sheetName);
  }
  
  spreadsheet.setActiveSheet(sheet);
}

/**
 * Show filter instructions
 */
function showFilterInstructions() {
  const ui = SpreadsheetApp.getUi();
  
  const instructions = `
ðŸ“ Filter Instructions:

When on a Summary Sheet:
â€¢ Use the filter options to quickly filter specific sections
â€¢ Each filter targets a specific data table in the summary
â€¢ Clear All Filters removes any active filters

When on Other Sheets:
â€¢ Select a filter option to choose which summary sheet to apply it to
â€¢ Use "Go to Summary Sheet" to navigate to a summary sheet first

Filter Types:
âš ï¸ LOE Exceeding Estimate - Shows epics where LOE > Story Point Estimate
ðŸ”„ Allocation Mismatches - Shows epics with allocation conflicts
ðŸ“Š All Epics - Shows all epics with sorting options

Tips:
â€¢ Filters are applied to specific sections only
â€¢ Each section maintains its own filter state
â€¢ Use Clear All Filters to reset the view
`;
  
  ui.alert('Filter Instructions', instructions, ui.ButtonSet.OK);
}

function getJiraValueStream(displayValueStream) {
  // For AIMM, we don't have a single value stream to search
  // The actual filtering happens in buildEpicJQL
  return displayValueStream;
}



// ===== WRAPPER FUNCTIONS =====

function startAnalysisWrapper(piNumber, selectedValueStreams) {
  try {
    analyzeSelectedValueStreams(piNumber, selectedValueStreams);
    return true;
  } catch (error) {
    console.error('Error in startAnalysisWrapper:', error);
    throw error;
  }
}

function startSummaryWrapper(piNumber, valueStream) {
  try {
    generateSummaryForValueStream(piNumber, valueStream);
    return true;
  } catch (error) {
    console.error('Error in startSummaryWrapper:', error);
    throw error;
  }
}

// ===== UI HELPER FUNCTIONS =====

function getCurrentPIFromSheets() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = spreadsheet.getSheets();
  
  const piPattern = /^PI (\d+)$/;
  let latestPI = null;
  
  sheets.forEach(sheet => {
    const match = sheet.getName().match(piPattern);
    if (match) {
      const piNumber = parseInt(match[1]);
      if (!latestPI || piNumber > latestPI) {
        latestPI = piNumber;
      }
    }
  });
  
  return latestPI ? latestPI.toString() : null;
}

// ===== SHEET MANIPULATION HELPERS =====

function menuRefreshFormulas() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    showProgress('Refreshing all formulas...');
    
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = spreadsheet.getSheets();
    let sheetsUpdated = 0;
    let totalFormulas = 0;
    
    sheets.forEach(sheet => {
      const sheetName = sheet.getName();
      showProgress(`Refreshing formulas in: ${sheetName}`);
      
      try {
        const dataRange = sheet.getDataRange();
        const formulas = dataRange.getFormulas();
        let sheetFormulas = 0;
        
        // Find and refresh all formulas
        for (let row = 0; row < formulas.length; row++) {
          for (let col = 0; col < formulas[row].length; col++) {
            if (formulas[row][col]) {
              const cell = sheet.getRange(row + 1, col + 1);
              cell.setFormula(formulas[row][col]);
              sheetFormulas++;
              totalFormulas++;
            }
          }
        }
        
        if (sheetFormulas > 0) {
          sheetsUpdated++;
          console.log(`${sheetName}: Refreshed ${sheetFormulas} formulas`);
        }
        
      } catch (error) {
        console.error(`Error refreshing ${sheetName}:`, error);
      }
    });
    
    SpreadsheetApp.flush(); // Force all pending changes
    closeProgress();
    
    ui.alert(
      'Formulas Refreshed',
      `Successfully refreshed ${totalFormulas} formulas across ${sheetsUpdated} sheets.`,
      ui.ButtonSet.OK
    );
    
  } catch (error) {
    closeProgress();
    ui.alert('Error', 'Error refreshing formulas: ' + error.toString(), ui.ButtonSet.OK);
  }
}

function safeCreateFilter(sheet, range) {
  try {
    if (!sheet || !range) return;
    
    safeRemoveFilter(sheet);
    Utilities.sleep(100);
    
    try {
      range.createFilter();
    } catch (e) {
      console.log('Could not create filter:', e.message);
      Utilities.sleep(500);
      try {
        range.createFilter();
      } catch (e2) {
        console.log('Filter creation failed after retry:', e2.message);
      }
    }
  } catch (error) {
    console.log('Error in safeCreateFilter:', error.message);
  }
}

// ===== CACHE FUNCTIONS =====
// NOTE: CACHE_EXPIRATION_MINUTES is defined in JiraPIAnalysis.gs

function getCachedData(cacheKey) {
  try {
    const cache = CacheService.getScriptCache();
    const cached = cache.get(cacheKey);
    if (cached) {
      return JSON.parse(cached);
    }
  } catch (e) {
    console.log('Cache read error:', e);
  }
  return null;
}

// ===== SCRUM TEAM DATA FUNCTIONS =====
// Note: The original updateScrumTeamData function is preserved in JiraPIAnalysis.gs
// The menu system uses enhanced versions that include automatic summary generation

/**
 * Get scrum teams from PI sheet
 */
function getScrumTeamsFromPISheet(piNumber) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const piSheet = spreadsheet.getSheetByName(`PI ${piNumber}`);
  
  if (!piSheet) {
    console.log(`No PI sheet found for PI ${piNumber}`);
    return [];
  }
  
  try {
    const dataRange = piSheet.getDataRange();
    const values = dataRange.getValues();
    
    // Check if sheet has enough rows
    if (!values || values.length <= 3) {
      console.log(`PI sheet ${piNumber} doesn't have enough rows. Expected headers at row 4.`);
      return [];
    }
    
    const headers = values[3]; // Header row at index 3
    
    // Validate headers
    if (!headers || !Array.isArray(headers) || headers.length === 0) {
      console.log(`PI sheet ${piNumber} has invalid header row at index 3`);
      return [];
    }
    
    // Helper function to extract value from field
    const getFieldValue = (field) => {
      if (!field) return 'Unassigned';
      
      // If it's already a string, check if it's a stringified object
      if (typeof field === 'string') {
        // Check if it looks like a stringified JIRA object {id=..., value=..., self=...}
        const objectPattern = /\{.*value=([^,}]+).*\}/;
        const match = field.match(objectPattern);
        if (match) {
          return match[1].trim();
        }
        return field;
      }
      
      // If it's a number, return as string
      if (typeof field === 'number') {
        return field.toString();
      }
      
      // If it's an object with value property
      if (field && typeof field === 'object') {
        if (field.value !== undefined && field.value !== null) {
          return field.value.toString();
        }
        if (field.name !== undefined && field.name !== null) {
          return field.name.toString();
        }
        
        // Check if it's a JIRA custom field object with all properties
        if (field.id !== undefined && field.self !== undefined && field.value !== undefined) {
          return field.value.toString();
        }
        
        // If it's still an object, check if it has a meaningful toString
        const stringified = field.toString();
        // Check if the toString gives us a JIRA object pattern
        const objectPattern = /\{.*value=([^,}]+).*\}/;
        const match = stringified.match(objectPattern);
        if (match) {
          return match[1].trim();
        }
        
        // Last resort
        if (stringified && stringified !== '[object Object]') {
          return stringified;
        }
      }
      
      return 'Unassigned';
    };
    
    // Find required column indices
    const scrumTeamIndex = headers.indexOf('Scrum Team');
    const issueTypeIndex = headers.indexOf('Issue Type');
    const valueStreamIndex = headers.indexOf('Value Stream');
    
    if (scrumTeamIndex === -1) {
      console.log(`PI sheet ${piNumber} doesn't have 'Scrum Team' column`);
      return [];
    }
    
    const teamMap = {};
    
    // Count items per team
    for (let i = 4; i < values.length; i++) {
      const row = values[i];
      if (!row || !row[0]) continue; // Skip empty rows
      
      // Extract the actual team name value
      const teamName = (scrumTeamIndex !== -1 && row[scrumTeamIndex]) ? 
                      getFieldValue(row[scrumTeamIndex]) : 'Unassigned';
      
      const issueType = (issueTypeIndex !== -1 && row[issueTypeIndex]) ? 
                       row[issueTypeIndex].toString() : '';
      
      const valueStream = (valueStreamIndex !== -1 && row[valueStreamIndex]) ? 
                         getFieldValue(row[valueStreamIndex]) : '';
      
      if (!teamMap[teamName]) {
        teamMap[teamName] = {
          name: teamName,
          count: 0,
          epicCount: 0,
          storyCount: 0,
          valueStream: valueStream
        };
      }
      
      teamMap[teamName].count++;
      
      if (issueType === 'Epic') {
        teamMap[teamName].epicCount++;
      } else {
        teamMap[teamName].storyCount++;
      }
    }
    
    return Object.values(teamMap)
      .filter(team => team.count > 0)
      .sort((a, b) => {
        if (a.name === 'Unassigned') return 1;
        if (b.name === 'Unassigned') return -1;
        return a.name.localeCompare(b.name);
      });
      
  } catch (error) {
    console.error(`Error reading scrum teams from PI ${piNumber}:`, error);
    return [];
  }
}


/**
 * Update data for a specific scrum team WITH automatic summary generation
 */
function updateScrumTeamDataWithSummary(piNumber, scrumTeamName) {
  const ui = SpreadsheetApp.getUi();
  const programIncrement = `PI ${piNumber}`;
  
  try {
    showProgress(`Updating data for ${scrumTeamName} in ${programIncrement}...`);
    
    // Check if PI sheet exists
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const piSheet = spreadsheet.getSheetByName(`PI ${piNumber}`);
    
    if (!piSheet) {
      closeProgress();
      ui.alert(`No data found for ${programIncrement}. Please run a full analysis first.`);
      return;
    }
    
    // For "Unassigned" team, just generate summary from existing data
    if (scrumTeamName === 'Unassigned') {
      showProgress('Generating summary from existing data...');
      
      // Read existing data
      const dataRange = piSheet.getDataRange();
      const values = dataRange.getValues();
      const headers = values[3];
      const allIssues = parsePISheetData(values, headers);
      
      // Filter for this team
      const teamIssues = allIssues.filter(issue => 
        (issue.scrumTeam || 'Unassigned') === scrumTeamName
      );
      
      if (teamIssues.length > 0) {
        // Generate summary
        createScrumTeamSummary(teamIssues, programIncrement, scrumTeamName);
        
        closeProgress();
        ui.alert(
          'Update Complete',
          `Summary generated for ${scrumTeamName} from existing data.\n\n` +
          `Epic count: ${teamIssues.filter(i => i.issueType === 'Epic').length}\n` +
          `Story count: ${teamIssues.filter(i => i.issueType !== 'Epic').length}`,
          ui.ButtonSet.OK
        );
      } else {
        closeProgress();
        ui.alert('No data found for Unassigned team.');
      }
      return;
    }
    
    // For assigned teams, fetch fresh data from JIRA
    showProgress(`Reading existing data for ${scrumTeamName}...`);
    
    // Read existing PI sheet data to get epic keys for this team
    const dataRange = piSheet.getDataRange();
    const values = dataRange.getValues();
    const headers = values[3];
    
    // Get column indices
    const keyCol = headers.indexOf('Key');
    const teamCol = headers.indexOf('Scrum Team');
    const issueTypeCol = headers.indexOf('Issue Type');
    
    if (keyCol === -1 || teamCol === -1 || issueTypeCol === -1) {
      closeProgress();
      ui.alert('PI sheet is missing required columns.');
      return;
    }
    
    // Extract epic keys for this team
    const teamEpicKeys = [];
    for (let i = 4; i < values.length; i++) {
      const row = values[i];
      if (!row[keyCol]) continue;
      
      // Check if this is an epic for our team
      if (row[issueTypeCol] === 'Epic' && row[teamCol] === scrumTeamName) {
        teamEpicKeys.push(row[keyCol]);
      }
    }
    
    console.log(`Found ${teamEpicKeys.length} epics for ${scrumTeamName}`);
    
    // BATCH fetch all epics at once instead of individually
    if (teamEpicKeys.length > 0 && scrumTeamName !== 'Unassigned') {
      showProgress(`Fetching ${teamEpicKeys.length} epics from JIRA...`);
      
      // Fetch all epics in one query
      const epicJql = `key in (${teamEpicKeys.join(',')}) AND status != "Closed"`;
      const epics = searchJiraIssues(epicJql);
      
      if (epics.length === 0) {
        closeProgress();
        ui.alert(`No active epics found for ${scrumTeamName} in JIRA.`);
        return;
      }
      
      showProgress(`Fetching child issues for ${epics.length} epics...`);
      
      // Fetch all children in batches
      const epicResults = [{
        valueStream: scrumTeamName,
        epics: epics,
        error: null
      }];
      
      const childrenMap = fetchChildIssuesInBatchesOptimized(epicResults);
      
      // Combine results
      const updatedIssues = [];
      epics.forEach(epic => {
        // Calculate LOE for epic
        const epicChildren = childrenMap[epic.key] || [];
        const loeEstimate = epicChildren.reduce((sum, child) => {
          if (child.issueType === 'Story' || child.issueType === 'Bug') {
            return sum + (parseFloat(child.storyPoints) || 0);
          }
          return sum;
        }, 0);
        
        // Add epic with calculated LOE
        updatedIssues.push({
          ...epic,
          loeEstimate: loeEstimate,
          analyzedValueStream: epic.valueStream
        });
        
        // Add children
        epicChildren.forEach(child => {
          updatedIssues.push({
            ...child,
            analyzedValueStream: epic.valueStream
          });
        });
      });
      
      showProgress('Updating sheet with fresh data...');
      updatePISheetForTeam(piSheet, scrumTeamName, updatedIssues);
      
      // Re-read the updated data
      showProgress('Generating summary...');
      const updatedDataRange = piSheet.getDataRange();
      const updatedValues = updatedDataRange.getValues();
      const updatedHeaders = updatedValues[3];
      const allUpdatedIssues = parsePISheetData(updatedValues, updatedHeaders);
      
      // Filter for this team
      const teamIssues = allUpdatedIssues.filter(issue => 
        (issue.scrumTeam || 'Unassigned') === scrumTeamName
      );
      
      // Generate summary
      createScrumTeamSummary(teamIssues, programIncrement, scrumTeamName);
      
      closeProgress();
      
      const epicCount = updatedIssues.filter(i => i.issueType === 'Epic').length;
      const storyCount = updatedIssues.filter(i => i.issueType !== 'Epic').length;
      
      ui.alert(
        'Update Complete',
        `Successfully updated ${scrumTeamName} in ${programIncrement}.\n\n` +
        `Updated from JIRA:\n` +
        `- Epics: ${epicCount}\n` +
        `- Stories/Tasks: ${storyCount}\n\n` +
        `Summary sheet has been created/updated.`,
        ui.ButtonSet.OK
      );
      
    } else {
      // No epics found for this team
      closeProgress();
      ui.alert(
        'No Data',
        `No epics found for ${scrumTeamName} in ${programIncrement}.\n\n` +
        `This team may not have any assigned work yet.`,
        ui.ButtonSet.OK
      );
    }
    
  } catch (error) {
    console.error('Error updating scrum team:', error);
    closeProgress();
    ui.alert('Error', 'Failed to update scrum team: ' + error.toString(), ui.ButtonSet.OK);
  }
}
/**
 * Batch update multiple scrum teams WITH automatic summary generation
 */
function UpdateScrumTeamsWithSummaries(piNumber, scrumTeams) {
  const ui = SpreadsheetApp.getUi();
  const programIncrement = `PI ${piNumber}`;
  
  // Sort teams alphabetically (case-insensitive) before processing
  const sortedTeams = scrumTeams.slice().sort(function(a, b) {
    return a.toLowerCase().localeCompare(b.toLowerCase());
  });
  console.log('Processing teams in alphabetical order: ' + sortedTeams.join(', '));
  
  try {
    showProgress(`Updating ${sortedTeams.length} scrum teams in ${programIncrement}...`);
    
    const results = [];
    
    sortedTeams.forEach((teamName, index) => {
      showProgress(`Updating ${teamName} (${index + 1}/${sortedTeams.length})...`);
      
      try {
        // For "Unassigned", don't try to fetch from JIRA
        const refreshFromJira = teamName !== 'Unassigned';
        
        // First update the team data
        const teamData = analyzeScrumTeamData(piNumber, teamName, refreshFromJira);
        
        // Then generate the summary
        if (teamData && teamData.issues.length > 0) {
          showProgress(`Creating summary for ${teamName}...`);
          createScrumTeamSummary(teamData.issues, programIncrement, teamName);
          
          results.push({
            team: teamName,
            success: true,
            epicCount: teamData.epicCount,
            storyCount: teamData.storyCount
          });
        } else {
          // No data found for team
          results.push({
            team: teamName,
            success: true,
            epicCount: 0,
            storyCount: 0,
            warning: 'No data found for this team'
          });
        }
      } catch (error) {
        console.error(`Error updating ${teamName}:`, error);
        results.push({
          team: teamName,
          success: false,
          error: error.message || error.toString()
        });
      }
    });
    
    closeProgress();
    
    // Show results
    const successCount = results.filter(r => r.success).length;
    const failureCount = results.filter(r => !r.success).length;
    const warningCount = results.filter(r => r.success && r.warning).length;
    
    let message = `Update Complete!\n\n`;
    message += `Successfully updated: ${successCount} teams\n`;
    
    if (successCount > warningCount) {
      message += `âœ“ Team summary sheets created/updated for teams with data\n`;
    }
    
    if (warningCount > 0) {
      message += `\nTeams with no data: ${warningCount}\n`;
      results.filter(r => r.success && r.warning).forEach(r => {
        message += `- ${r.team}: ${r.warning}\n`;
      });
    }
    
    if (failureCount > 0) {
      message += `\nFailed: ${failureCount} teams\n`;
      message += 'Failed teams:\n';
      results.filter(r => !r.success).forEach(r => {
        message += `- ${r.team}: ${r.error}\n`;
      });
    }
    
    ui.alert('Batch Update Results', message, ui.ButtonSet.OK);
    
  } catch (error) {
    console.error('Error in batch update:', error);
    closeProgress();
    ui.alert('Error', 'Failed to complete batch update: ' + error.toString(), ui.ButtonSet.OK);
  }
}

// Alias for backward compatibility
function batchUpdateScrumTeams(piNumber, scrumTeams) {
  UpdateScrumTeamsWithSummaries(piNumber, scrumTeams);
}

/**
 * Generate summaries for selected scrum teams
 */
function generateScrumTeamSummaries(piNumber, selectedTeams) {
  const ui = SpreadsheetApp.getUi();
  const programIncrement = `PI ${piNumber}`;
  
  try {
    showProgress(`Generating summaries for ${selectedTeams.length} scrum teams...`);
    
    // Get PI sheet data
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const piSheet = spreadsheet.getSheetByName(`PI ${piNumber}`);
    
    if (!piSheet) {
      throw new Error('PI sheet not found');
    }
    
    // Read and parse data
    const dataRange = piSheet.getDataRange();
    const values = dataRange.getValues();
    const headers = values[3];
    
    const issues = parsePISheetData(values, headers);
    
    // Generate summary for each selected team
    selectedTeams.forEach((teamName, index) => {
      showProgress(`Creating summary for ${teamName} (${index + 1}/${selectedTeams.length})...`);
      
      const teamIssues = issues.filter(issue => issue.scrumTeam === teamName);
      
      if (teamIssues.length > 0) {
        createScrumTeamSummary(teamIssues, programIncrement, teamName);
      }
    });
    
    closeProgress();
    
    ui.alert(
      'Summaries Complete',
      `Generated summary sheets for ${selectedTeams.length} scrum teams in ${programIncrement}.`,
      ui.ButtonSet.OK
    );
    
  } catch (error) {
    console.error('Error generating summaries:', error);
    closeProgress();
    ui.alert('Error', 'Failed to generate summaries: ' + error.toString(), ui.ButtonSet.OK);
  }
}

// ===== HELPER FUNCTIONS =====
/**
 * Get column name for a field
 */
function getColumnNameForField(field) {
  const fieldToColumnMap = {
    'summary': 'Summary',
    'status': 'Status',
    'storyPoints': 'Story Points',
    'storyPointEstimate': 'Story Point Estimate',
    'valueStream': 'Value Stream',
    'org': 'Org',
    'piCommitment': 'PI Commitment',
    'programIncrement': 'Program Increment',
    'epicLink': 'Epic Link',
    'scrumTeam': 'Scrum Team',
    'piTargetIteration': 'PI Target Iteration',
    'iterationStart': 'Iteration Start',
    'iterationEnd': 'Iteration End',
    'allocation': 'Allocation',
    'portfolioInitiative': 'Portfolio Initiative',
    'programInitiative': 'Program Initiative',
    'featurePoints': 'Feature Points',
    'rag': 'RAG',
    'ragNote': 'RAG Note',
    'dependsOnValuestream': 'Depends on Valuestream',
    'costOfDelay': 'Cost of Delay'
  };
  
  return fieldToColumnMap[field] || field;
}

/**
 * Discover all scrum teams from JIRA
 */
function discoverScrumTeamsFromJira() {
  try {
    console.log('Discovering all scrum teams in JIRA across ALL projects...');
    
    // Search for epics with scrum teams
    const jql = `issuetype = Epic AND cf[10040] is not EMPTY ORDER BY created DESC`;
    const url = `${JIRA_CONFIG.baseUrl}/rest/api/3/search/jql`;
    
    const scrumTeams = new Set();
    let startAt = 0;
    const maxResults = 100;
    let totalProcessed = 0;
    
    while (totalProcessed < 1000) {
      const payload = {
        jql: jql,
        startAt: startAt,
        maxResults: maxResults,
        fields: ['customfield_10040'] // Scrum Team field
      };
      
      const response = makeJiraRequest(url, 'POST', payload);
      
      if (response && response.issues) {
        response.issues.forEach(issue => {
          const scrumTeam = issue.fields.customfield_10040;
          
          if (scrumTeam) {
            const teamValue = scrumTeam.value || scrumTeam;
            if (teamValue && teamValue.toString().trim()) {
              scrumTeams.add(teamValue.toString().trim());
            }
          }
        });
        
        totalProcessed += response.issues.length;
        
        if (response.issues.length < maxResults || totalProcessed >= response.total) {
          break;
        }
        startAt += maxResults;
      } else {
        break;
      }
    }
    
    const scrumTeamArray = Array.from(scrumTeams).sort();
    console.log(`Found ${scrumTeamArray.length} scrum teams across all projects:`, scrumTeamArray);
    
    return scrumTeamArray;
    
  } catch (error) {
    console.error('Error discovering scrum teams:', error);
    throw error;
  }
}

// ===== IDENTIFY SHEETS WITH FORMULAS =====

/**
 * Utility function to identify which sheets have formulas
 * Run this to understand your spreadsheet structure
 */
function identifySheetsWithFormulas() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = spreadsheet.getSheets();
  const report = [];
  
  sheets.forEach(sheet => {
    const sheetName = sheet.getName();
    const range = sheet.getDataRange();
    const formulas = range.getFormulas();
    let formulaCount = 0;
    const sampleFormulas = [];
    
    for (let row = 0; row < formulas.length; row++) {
      for (let col = 0; col < formulas[row].length; col++) {
        if (formulas[row][col]) {
          formulaCount++;
          if (sampleFormulas.length < 3) {
            sampleFormulas.push({
              cell: `${String.fromCharCode(65 + col)}${row + 1}`,
              formula: formulas[row][col]
            });
          }
        }
      }
    }
    
    if (formulaCount > 0) {
      report.push({
        sheetName: sheetName,
        formulaCount: formulaCount,
        samples: sampleFormulas
      });
    }
  });
  
  console.log('Sheets with formulas:');
  report.forEach(item => {
    console.log(`\nSheet: ${item.sheetName}`);
    console.log(`Formula count: ${item.formulaCount}`);
    console.log('Sample formulas:');
    item.samples.forEach(sample => {
      console.log(`  ${sample.cell}: ${sample.formula}`);
    });
  });
  
  return report;
}

function cleanCurrentSheetData() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSheet();
  
  const response = ui.alert(
    'Clean Sheet Data',
    'This will clean up any stringified JIRA objects in the current sheet.\n\nContinue?',
    ui.ButtonSet.YES_NO
  );
  
  if (response === ui.Button.YES) {
    showProgress('Cleaning sheet data...');
    const cleanedCount = postProcessSheetData(sheet);
    closeProgress();
    
    if (cleanedCount > 0) {
      ui.alert('Success', `Cleaned ${cleanedCount} cells with stringified objects.`, ui.ButtonSet.OK);
    } else {
      ui.alert('No Changes', 'No stringified objects found in this sheet.', ui.ButtonSet.OK);
    }
  }
}

// =============================================================================
// HELPER FUNCTIONS FOR REPORT DESTINATIONS
// =============================================================================

/**
 * Copy a sheet by name from source to target if it exists.
 * Skips if target already has a sheet with the same name.
 * @param {Spreadsheet} source - Source spreadsheet
 * @param {Spreadsheet} target - Target spreadsheet
 * @param {string} sheetName - Name of the sheet to copy
 * @private
 */
function copySheetIfExists_(source, target, sheetName) {
  var sheet = source.getSheetByName(sheetName);
  if (sheet) {
    try {
      // Check if target already has this sheet
      var existing = target.getSheetByName(sheetName);
      if (existing) {
        console.log('"' + sheetName + '" already exists in target, skipping copy');
        return;
      }
      sheet.copyTo(target).setName(sheetName);
      console.log('Copied "' + sheetName + '" to target spreadsheet');
    } catch (e) {
      console.warn('Could not copy "' + sheetName + '":', e);
    }
  }
}

/**
 * Count epics in a PI sheet within a spreadsheet.
 * @param {Spreadsheet} spreadsheet - Target spreadsheet
 * @param {string} piNumber - PI number
 * @returns {number} Epic count
 * @private
 */
function countEpicsInSpreadsheet_(spreadsheet, piNumber) {
  try {
    var piSheet = spreadsheet.getSheetByName('PI ' + piNumber);
    if (!piSheet) return 0;
    
    var data = piSheet.getDataRange().getValues();
    if (data.length <= 4) return 0;
    
    var headers = data[3];
    var typeCol = headers.indexOf('Issue Type');
    if (typeCol === -1) return 0;
    
    var count = 0;
    for (var i = 4; i < data.length; i++) {
      if (data[i][typeCol] === 'Epic') count++;
    }
    return count;
  } catch (e) {
    return 0;
  }
}