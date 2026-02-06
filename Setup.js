/**
 * Setup.gs - Credential Setup and Saved Reports Management
 * =========================================================
 * 
 * Manages sensitive credentials and saved report URLs using 
 * Google Apps Script Properties.
 * 
 * FIRST TIME SETUP:
 * 1. Open the spreadsheet
 * 2. Run "JIRA Analysis > Setup > Configure JIRA Credentials"
 * 3. Enter your JIRA email and API token
 * 
 * TO GET A JIRA API TOKEN:
 * 1. Go to https://id.atlassian.com/manage-profile/security/api-tokens
 * 2. Click "Create API token"
 * 3. Copy the token (you won't see it again)
 * 
 * @fileoverview Credential and saved reports management
 * @version 1.0.0
 */

// Property keys for Script Properties storage
const PROP_KEYS = {
  JIRA_EMAIL: 'JIRA_EMAIL',
  JIRA_API_TOKEN: 'JIRA_API_TOKEN',
  JIRA_BASE_URL: 'JIRA_BASE_URL',
  SETUP_COMPLETE: 'SETUP_COMPLETE',
  SAVED_REPORTS: 'SAVED_REPORTS'
};

// ============================================================================
// CREDENTIAL MANAGEMENT
// ============================================================================

/**
 * Check if JIRA credentials are configured
 * @returns {boolean} True if credentials exist
 */
function isJiraConfigured() {
  const props = PropertiesService.getScriptProperties();
  return !!(props.getProperty(PROP_KEYS.JIRA_EMAIL) && 
            props.getProperty(PROP_KEYS.JIRA_API_TOKEN));
}

/**
 * Get JIRA credentials from Script Properties
 * @returns {Object|null} Credentials object or null if not configured
 */
function getJiraCredentials() {
  const props = PropertiesService.getScriptProperties();
  
  const email = props.getProperty(PROP_KEYS.JIRA_EMAIL);
  const token = props.getProperty(PROP_KEYS.JIRA_API_TOKEN);
  const baseUrl = props.getProperty(PROP_KEYS.JIRA_BASE_URL) || 'https://modmedrnd.atlassian.net';
  
  if (!email || !token) {
    return null;
  }
  
  return { email, apiToken: token, baseUrl };
}

/**
 * Save JIRA credentials to Script Properties
 * @param {string} email - JIRA email address
 * @param {string} apiToken - JIRA API token
 * @param {string} [baseUrl] - Optional JIRA base URL
 * @returns {boolean} True if saved successfully
 */
function saveJiraCredentials(email, apiToken, baseUrl) {
  try {
    const props = PropertiesService.getScriptProperties();
    
    props.setProperty(PROP_KEYS.JIRA_EMAIL, email.trim());
    props.setProperty(PROP_KEYS.JIRA_API_TOKEN, apiToken.trim());
    
    if (baseUrl) {
      props.setProperty(PROP_KEYS.JIRA_BASE_URL, baseUrl.trim().replace(/\/$/, ''));
    }
    
    props.setProperty(PROP_KEYS.SETUP_COMPLETE, 'true');
    
    console.log('✅ JIRA credentials saved successfully');
    return true;
  } catch (error) {
    console.error('Error saving credentials:', error);
    return false;
  }
}

/**
 * Clear all JIRA credentials
 * @returns {boolean} True if cleared successfully
 */
function clearJiraCredentials() {
  try {
    const props = PropertiesService.getScriptProperties();
    props.deleteProperty(PROP_KEYS.JIRA_EMAIL);
    props.deleteProperty(PROP_KEYS.JIRA_API_TOKEN);
    props.deleteProperty(PROP_KEYS.JIRA_BASE_URL);
    props.deleteProperty(PROP_KEYS.SETUP_COMPLETE);
    console.log('✅ JIRA credentials cleared');
    return true;
  } catch (error) {
    console.error('Error clearing credentials:', error);
    return false;
  }
}

// ============================================================================
// JIRA_CONFIG COMPATIBILITY LAYER
// ============================================================================

/**
 * Get JIRA config object (for backward compatibility with existing code)
 * This replaces the hardcoded JIRA_CONFIG object
 * @returns {Object} JIRA configuration
 */
function getJiraConfig() {
  const creds = getJiraCredentials();
  
  if (!creds) {
    console.warn('⚠️ JIRA credentials not configured. Run Setup > Configure JIRA Credentials');
    return {
      baseUrl: 'https://modmedrnd.atlassian.net',
      email: '',
      apiToken: ''
    };
  }
  
  return {
    baseUrl: creds.baseUrl,
    email: creds.email,
    apiToken: creds.apiToken
  };
}

// ============================================================================
// SAVED REPORTS MANAGEMENT
// ============================================================================

/**
 * Get all saved reports
 * @returns {Array} Array of saved report objects [{label, url, dateAdded}, ...]
 */
function getSavedReports() {
  try {
    const props = PropertiesService.getScriptProperties();
    const savedReportsJson = props.getProperty(PROP_KEYS.SAVED_REPORTS);
    
    if (!savedReportsJson) {
      return [];
    }
    
    const reports = JSON.parse(savedReportsJson);
    
    // Sort by date added (newest first)
    reports.sort((a, b) => {
      const dateA = a.dateAdded ? new Date(a.dateAdded) : new Date(0);
      const dateB = b.dateAdded ? new Date(b.dateAdded) : new Date(0);
      return dateB - dateA;
    });
    
    return reports;
  } catch (error) {
    console.error('Error getting saved reports:', error);
    return [];
  }
}

/**
 * Save a report URL for future use
 * @param {string} label - User-friendly label (e.g., "PI 14 - MMPM")
 * @param {string} url - Google Sheets URL
 * @returns {boolean} True if saved successfully
 */
function saveReport(label, url) {
  try {
    const props = PropertiesService.getScriptProperties();
    
    let reports = getSavedReports();
    
    // Check if URL already exists - update label if so
    const existingIndex = reports.findIndex(r => r.url === url);
    if (existingIndex >= 0) {
      reports[existingIndex].label = label;
      reports[existingIndex].dateAdded = new Date().toISOString();
      console.log(`Updated existing saved report: ${label}`);
    } else {
      reports.push({
        label: label,
        url: url,
        dateAdded: new Date().toISOString()
      });
      console.log(`Saved new report: ${label}`);
    }
    
    props.setProperty(PROP_KEYS.SAVED_REPORTS, JSON.stringify(reports));
    return true;
  } catch (error) {
    console.error('Error saving report:', error);
    return false;
  }
}

/**
 * Delete a saved report by URL
 * @param {string} url - URL of the report to delete
 * @returns {boolean} True if deleted successfully
 */
function deleteSavedReport(url) {
  try {
    const props = PropertiesService.getScriptProperties();
    
    let reports = getSavedReports();
    const originalLength = reports.length;
    reports = reports.filter(r => r.url !== url);
    
    if (reports.length === originalLength) {
      console.log('Report not found for deletion');
      return false;
    }
    
    props.setProperty(PROP_KEYS.SAVED_REPORTS, JSON.stringify(reports));
    console.log('Report deleted successfully');
    return true;
  } catch (error) {
    console.error('Error deleting report:', error);
    return false;
  }
}

/**
 * Clear all saved reports
 * @returns {boolean} True if cleared successfully
 */
function clearAllSavedReports() {
  try {
    const props = PropertiesService.getScriptProperties();
    props.deleteProperty(PROP_KEYS.SAVED_REPORTS);
    console.log('All saved reports cleared');
    return true;
  } catch (error) {
    console.error('Error clearing saved reports:', error);
    return false;
  }
}

// ============================================================================
// SETUP DIALOG
// ============================================================================

/**
 * Show the JIRA credential setup dialog
 */
function showCredentialSetupDialog() {
  const currentCreds = getJiraCredentials();
  const hasExisting = !!currentCreds;
  const defaultBaseUrl = 'https://modmedrnd.atlassian.net';
  
  const html = `
    <!DOCTYPE html>
    <html>
      <head>
        <base target="_top">
        <style>
          body { font-family: 'Roboto', Arial, sans-serif; padding: 20px; }
          .form-group { margin-bottom: 15px; }
          label { display: block; font-weight: 500; margin-bottom: 5px; color: #1B365D; }
          input { width: 100%; padding: 10px; border: 1px solid #ddd; border-radius: 4px; font-size: 14px; box-sizing: border-box; }
          input:focus { outline: none; border-color: #6B3FA0; box-shadow: 0 0 0 2px #E8DEF8; }
          .button-group { margin-top: 20px; display: flex; gap: 10px; justify-content: flex-end; }
          button { padding: 10px 20px; border: none; border-radius: 4px; font-size: 14px; cursor: pointer; }
          .primary { background-color: #1B365D; color: white; }
          .primary:hover { background-color: #152a4a; }
          .secondary { background-color: #f1f3f4; color: #5f6368; }
          .danger { background-color: #C62828; color: white; }
          .info-box { background-color: #e3f2fd; border-left: 4px solid #1B365D; padding: 12px; margin-bottom: 20px; font-size: 13px; }
          .success-box { background-color: #C8E6C9; border-left: 4px solid #2E7D32; padding: 12px; margin-bottom: 20px; font-size: 13px; }
          .help-text { font-size: 12px; color: #666; margin-top: 5px; }
          .help-text a { color: #6B3FA0; }
          #status { margin-top: 15px; text-align: center; padding: 10px; border-radius: 4px; }
          .error { background-color: #FFCDD2; color: #C62828; }
          .success { background-color: #C8E6C9; color: #2E7D32; }
        </style>
      </head>
      <body>
        <h3 style="color: #1B365D; border-bottom: 3px solid #FFC72C; padding-bottom: 10px;">Configure JIRA Credentials</h3>
        
        ${hasExisting ? `
          <div class="success-box">
            ✓ Credentials configured for: <strong>${currentCreds.email}</strong>
          </div>
        ` : `
          <div class="info-box">
            Enter your JIRA credentials below. These are stored securely in Script Properties and never appear in the code.
          </div>
        `}
        
        <div class="form-group">
          <label for="email">JIRA Email Address</label>
          <input type="text" id="email" placeholder="your.email@modmed.com" 
                 value="${hasExisting ? currentCreds.email : ''}">
        </div>
        
        <div class="form-group">
          <label for="token">JIRA API Token</label>
          <input type="password" id="token" placeholder="Enter your API token">
          <div class="help-text">
            <a href="https://id.atlassian.com/manage-profile/security/api-tokens" target="_blank">
              Get an API token from Atlassian →
            </a>
          </div>
        </div>
        
        <div class="form-group">
          <label for="baseUrl">JIRA Base URL</label>
          <input type="text" id="baseUrl" 
                 value="${hasExisting && currentCreds.baseUrl ? currentCreds.baseUrl : defaultBaseUrl}">
          <div class="help-text">Only change if using a different JIRA instance</div>
        </div>
        
        <div id="status"></div>
        
        <div class="button-group">
          ${hasExisting ? '<button class="danger" onclick="clearCreds()">Clear Credentials</button>' : ''}
          <button class="secondary" onclick="google.script.host.close()">Cancel</button>
          <button class="primary" onclick="saveCreds()">Save & Test</button>
        </div>
        
        <script>
          function saveCreds() {
            const email = document.getElementById('email').value.trim();
            const token = document.getElementById('token').value.trim();
            const baseUrl = document.getElementById('baseUrl').value.trim();
            const status = document.getElementById('status');
            
            if (!email || !token) {
              status.className = 'error';
              status.innerText = 'Please enter both email and API token.';
              return;
            }
            
            status.className = '';
            status.innerText = 'Saving and testing connection...';
            
            google.script.run
              .withSuccessHandler(function(result) {
                if (result) {
                  status.className = 'success';
                  status.innerText = '✓ Saved successfully!';
                  setTimeout(() => google.script.host.close(), 1500);
                } else {
                  status.className = 'error';
                  status.innerText = 'Failed to save. Please try again.';
                }
              })
              .withFailureHandler(function(err) {
                status.className = 'error';
                status.innerText = 'Error: ' + err.message;
              })
              .saveJiraCredentials(email, token, baseUrl);
          }
          
          function clearCreds() {
            if (!confirm('Clear your JIRA credentials? You will need to reconfigure them to use the system.')) return;
            
            google.script.run
              .withSuccessHandler(function() {
                document.getElementById('status').className = 'success';
                document.getElementById('status').innerText = 'Credentials cleared.';
                setTimeout(() => google.script.host.close(), 1000);
              })
              .clearJiraCredentials();
          }
        </script>
      </body>
    </html>
  `;
  
  const htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(450)
    .setHeight(hasExisting ? 480 : 420);
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'JIRA Setup');
}

/**
 * Test JIRA connection with stored credentials
 * @returns {Object} Test result with success boolean and message
 */
function testJiraConnection() {
  const creds = getJiraCredentials();
  
  if (!creds) {
    return { success: false, message: 'No credentials configured. Run Setup > Configure JIRA Credentials first.' };
  }
  
  try {
    const url = `${creds.baseUrl}/rest/api/3/myself`;
    const response = UrlFetchApp.fetch(url, {
      method: 'GET',
      headers: {
        'Authorization': 'Basic ' + Utilities.base64Encode(`${creds.email}:${creds.apiToken}`),
        'Accept': 'application/json'
      },
      muteHttpExceptions: true
    });
    
    const code = response.getResponseCode();
    
    if (code === 200) {
      const user = JSON.parse(response.getContentText());
      return { success: true, message: `Connected as: ${user.displayName} (${user.emailAddress})` };
    } else if (code === 401) {
      return { success: false, message: 'Authentication failed. Check your email and API token.' };
    } else {
      return { success: false, message: `Connection failed (HTTP ${code}). Check your JIRA URL.` };
    }
  } catch (error) {
    return { success: false, message: `Error: ${error.message}` };
  }
}

/**
 * Menu handler for testing connection
 */
function menuTestJiraConnection() {
  const ui = SpreadsheetApp.getUi();
  
  if (!isJiraConfigured()) {
    const response = ui.alert('Setup Required', 
      'JIRA credentials not configured. Set them up now?', 
      ui.ButtonSet.YES_NO);
    if (response === ui.Button.YES) {
      showCredentialSetupDialog();
    }
    return;
  }
  
  ui.alert('Testing...', 'Testing JIRA connection...', ui.ButtonSet.OK);
  
  const result = testJiraConnection();
  ui.alert(
    result.success ? '✓ Connection Successful' : '✗ Connection Failed', 
    result.message, 
    ui.ButtonSet.OK
  );
}

// ============================================================================
// REPORT LOG FUNCTIONS
// ============================================================================

/**
 * Get or create the Report Log sheet
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} The Report Log sheet
 */
function getOrCreateReportLogSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = 'Report Log';
  let logSheet = spreadsheet.getSheetByName(sheetName);
  
  if (!logSheet) {
    logSheet = spreadsheet.insertSheet(sheetName);
    
    // Setup headers
    const headers = [
      'Generated Date', 'PI', 'Value Stream', 
      'Report Name', 'Spreadsheet URL', 'Spreadsheet ID', 
      'Epic Count', 'Status'
    ];
    
    logSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    logSheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#9b7bb8')
      .setFontColor('#ffffff')
      .setHorizontalAlignment('center');
    
    // Set column widths
    const widths = [150, 60, 120, 250, 350, 200, 80, 100];
    widths.forEach((w, i) => logSheet.setColumnWidth(i + 1, w));
    
    logSheet.setFrozenRows(1);
  }
  
  return logSheet;
}

/**
 * Log a report to the Report Log sheet
 * @param {Object} reportInfo - Report metadata
 */
function logReport(reportInfo) {
  const logSheet = getOrCreateReportLogSheet();
  const lastRow = Math.max(logSheet.getLastRow(), 1);
  
  const logData = [
    new Date().toLocaleString(),
    reportInfo.piNumber ? `PI ${reportInfo.piNumber}` : reportInfo.pi || '',
    reportInfo.valueStream || '',
    reportInfo.reportName || '',
    reportInfo.spreadsheetUrl || '',
    reportInfo.spreadsheetId || '',
    reportInfo.epicCount || 0,
    reportInfo.status || 'Success'
  ];
  
  logSheet.getRange(lastRow + 1, 1, 1, logData.length).setValues([logData]);
  
  // Make URL clickable
  if (reportInfo.spreadsheetUrl) {
    const urlCell = logSheet.getRange(lastRow + 1, 5);
    urlCell.setFormula(`=HYPERLINK("${reportInfo.spreadsheetUrl}", "Open Report")`);
    urlCell.setFontColor('#6B3FA0');
  }
  
  console.log('📝 Report logged:', reportInfo.reportName);
}

/**
 * View the Report Log sheet
 */
function viewReportLog() {
  const logSheet = getOrCreateReportLogSheet();
  SpreadsheetApp.setActiveSheet(logSheet);
}

/**
 * Open a report from the log
 */
function openReportFromLog() {
  const ui = SpreadsheetApp.getUi();
  const logSheet = getOrCreateReportLogSheet();
  const data = logSheet.getDataRange().getValues();
  
  if (data.length <= 1) {
    ui.alert('No Reports', 'No reports have been generated yet.', ui.ButtonSet.OK);
    return;
  }
  
  // Build list of recent reports (last 10)
  const recentReports = [];
  for (let i = Math.min(data.length - 1, 11); i >= 1; i--) {
    const row = data[i];
    if (row[5]) { // Has spreadsheet ID
      recentReports.push({
        index: i,
        name: `${row[1]} - ${row[2]} (${row[0]})`,
        url: row[4],
        id: row[5]
      });
    }
  }
  
  if (recentReports.length === 0) {
    ui.alert('No Reports', 'No reports with valid URLs found.', ui.ButtonSet.OK);
    return;
  }
  
  // Show selection dialog
  const reportList = recentReports.map((r, i) => `${i + 1}. ${r.name}`).join('\n');
  const response = ui.prompt('Open Report',
    `Enter the number of the report to open:\n\n${reportList}`,
    ui.ButtonSet.OK_CANCEL);
  
  if (response.getSelectedButton() !== ui.Button.OK) return;
  
  const selection = parseInt(response.getResponseText()) - 1;
  if (isNaN(selection) || selection < 0 || selection >= recentReports.length) {
    ui.alert('Invalid Selection', 'Please enter a valid number.', ui.ButtonSet.OK);
    return;
  }
  
  const selectedReport = recentReports[selection];
  
  // Open the report in a new tab
  const html = HtmlService.createHtmlOutput(
    `<script>window.open("${selectedReport.url}", "_blank"); google.script.host.close();</script>`
  ).setWidth(1).setHeight(1);
  
  ui.showModalDialog(html, 'Opening Report...');
}