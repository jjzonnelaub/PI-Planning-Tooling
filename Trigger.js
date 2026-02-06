/**
 * ========================================
 * TRIGGER FILE - 15 MINUTE UPDATES
 * ========================================
 * 
 * This trigger runs every 15 minutes and uses smart caching:
 * - Refreshes from JIRA at :00 and :30 (twice per hour)
 * - Uses cache at :15 and :45 (fast, no API calls)
 * 
 * Setup Instructions:
 * 1. Copy all functions below to your trigger.gs file
 * 2. In Apps Script, go to Triggers (clock icon)
 * 3. Create/Edit trigger:
 *    - Function: hourlyUpdate
 *    - Event source: Time-driven
 *    - Type: Minute timer
 *    - Interval: Every 15 minutes
 * 4. Save and authorize
 */

/**
 * Main trigger function - runs every 15 minutes
 */
function hourlyUpdate() {
  // ===== CONFIGURATION =====
  const PI_NUMBER = '13';  // Update this to your current PI
  const VALUE_STREAMS = [
    'AIMM',
    'EMA Clinical', 
    'EMA RAC',
    'MMPM',
    'RCM'
  ];
  
  // ===== EXECUTION =====
  try {
    const now = new Date();
    const timeStr = now.toLocaleTimeString();
    const minutes = now.getMinutes();
    
    console.log(`🕐 Starting 15-minute update for PI ${PI_NUMBER} at ${timeStr}...`);
    console.log(`Updating value streams: ${VALUE_STREAMS.join(', ')}`);
    
    // Determine if this is a refresh cycle
    const isRefreshCycle = (minutes === 0 || minutes === 30);
    if (isRefreshCycle) {
      console.log('🔄 Refresh cycle - will fetch fresh data from JIRA');
    } else {
      console.log('⚡ Cache cycle - will use cached data if available');
    }
    
    // Enable trigger-safe mode
    setupTriggerSafeMode();
    
    // Track execution time
    const startTime = new Date();
    
    // Call the main analysis function
    analyzeSelectedValueStreams(PI_NUMBER, VALUE_STREAMS);
    
    // Calculate duration
    const duration = ((new Date() - startTime) / 1000).toFixed(1);
    
    console.log('========================================');
    console.log(`✅ Update completed successfully in ${duration}s`);
    console.log(`Completed at: ${new Date().toLocaleTimeString()}`);
    console.log('========================================');
    
  } catch (error) {
    console.error('========================================');
    console.error('❌ ERROR in 15-minute update:', error);
    console.error('Stack trace:', error.stack);
    console.error('========================================');
    
    // Send email notification on failure
    sendErrorEmail(error);
  }
}

/**
 * Makes UI functions trigger-safe with smart caching
 * Option A: Refresh at :00 and :30, cache at :15 and :45
 */
function setupTriggerSafeMode() {
  // Override the global showProgress function
  globalThis.showProgress = function(message) {
    console.log('[PROGRESS]', message);
  };
  
  // Override the global closeProgress function  
  globalThis.closeProgress = function() {
    console.log('[PROGRESS] Complete');
  };
  
  // Create a mock UI object for SpreadsheetApp.getUi()
  const mockUi = {
    alert: function(title, message, buttons) {
      console.log('[UI ALERT]', title);
      if (message) console.log('[UI ALERT]', message);
      
      // Smart cache decision based on time
      if (title.includes('Cache Option')) {
        const now = new Date();
        const minutes = now.getMinutes();
        
        // ⭐ OPTION A: Refresh at :00 and :30 (twice per hour)
        if (minutes === 0 || minutes === 30) {
          console.log('🔄 Refresh window - forcing fresh JIRA data');
          return mockUi.Button.NO; // NO = don't use cache, fetch fresh
        } else {
          console.log('⚡ Cache window - using cached data for speed');
          return mockUi.Button.YES; // YES = use cache
        }
      }
      
      // For other alerts, return OK
      return mockUi.Button.OK;
    },
    Button: { NO: 'NO', YES: 'YES', OK: 'OK' },
    ButtonSet: { YES_NO: 'YES_NO', OK: 'OK' }
  };
  
  // Override SpreadsheetApp.getUi()
  SpreadsheetApp.getUi = function() {
    return mockUi;
  };
  
  console.log('✅ Trigger-safe mode enabled (Option A: refresh at :00 and :30)');
}

/**
 * Send email notification on error
 */
function sendErrorEmail(error) {
  try {
    const userEmail = Session.getEffectiveUser().getEmail();
    const subject = 'JIRA Dashboard - 15-Minute Update Failed';
    const body = `The 15-minute JIRA integration update failed with the following error:

ERROR:
${error.toString()}

STACK TRACE:
${error.stack || 'No stack trace available'}

TIMESTAMP:
${new Date().toLocaleString()}

TRIGGER SCHEDULE:
Every 15 minutes with smart caching
- Refreshes from JIRA at :00 and :30
- Uses cache at :15 and :45

Please check the Apps Script execution logs for more details:
Extensions > Apps Script > Executions

---
This is an automated notification from your JIRA PI Planning Dashboard.`;
    
    MailApp.sendEmail(userEmail, subject, body);
    console.log(`✅ Error notification sent to ${userEmail}`);
  } catch (emailError) {
    console.error('❌ Failed to send error email:', emailError);
  }
}

/**
 * Optional: Send success notification email
 * Uncomment the call in hourlyUpdate() if you want success notifications
 * (Warning: This will send 96 emails per day!)
 */
function sendSuccessEmail(piNumber, valueStreams, duration, usedCache) {
  try {
    const userEmail = Session.getEffectiveUser().getEmail();
    const subject = `JIRA Dashboard - PI ${piNumber} Update Complete`;
    const cacheStatus = usedCache ? 'Used cached data (fast)' : 'Fetched fresh data from JIRA';
    
    const body = `The 15-minute JIRA integration update completed successfully.

PI: ${piNumber}
VALUE STREAMS: ${valueStreams.join(', ')}
DURATION: ${duration}s
CACHE STATUS: ${cacheStatus}
TIMESTAMP: ${new Date().toLocaleString()}

Your dashboard has been updated with the latest data.

---
This is an automated notification from your JIRA PI Planning Dashboard.`;
    
    MailApp.sendEmail(userEmail, subject, body);
    console.log(`✅ Success notification sent to ${userEmail}`);
  } catch (emailError) {
    console.error('❌ Failed to send success email:', emailError);
  }
}

/**
 * ========================================
 * TESTING & DIAGNOSTIC FUNCTIONS
 * ========================================
 */

/**
 * Test function - run this manually to verify trigger setup
 */
function testFifteenMinuteUpdate() {
  console.log('=== MANUAL TEST RUN ===');
  console.log('This simulates the trigger execution without consuming trigger quota');
  console.log('');
  
  hourlyUpdate();
  
  console.log('');
  console.log('=== TEST COMPLETE ===');
  console.log('Check the logs above for any errors');
  console.log('If successful, your trigger is properly configured!');
}

/**
 * Show current cache status and schedule
 */
function showCacheSchedule() {
  const now = new Date();
  const currentMinute = now.getMinutes();
  
  console.log('========================================');
  console.log('CACHE SCHEDULE (Option A)');
  console.log('========================================');
  console.log('Current time:', now.toLocaleTimeString());
  console.log('Current minute:', currentMinute);
  console.log('');
  console.log('REFRESH WINDOWS (fetch from JIRA):');
  console.log('  - XX:00 - Beginning of each hour');
  console.log('  - XX:30 - Middle of each hour');
  console.log('');
  console.log('CACHE WINDOWS (use cached data):');
  console.log('  - XX:15 - 15 minutes past');
  console.log('  - XX:45 - 45 minutes past');
  console.log('');
  
  if (currentMinute === 0 || currentMinute === 30) {
    console.log('🔄 Current status: REFRESH WINDOW');
    console.log('Next run will fetch fresh data from JIRA');
  } else {
    console.log('⚡ Current status: CACHE WINDOW');
    console.log('Next run will use cached data');
  }
  
  console.log('========================================');
}

/**
 * Clear cache manually (useful for testing or forcing refresh)
 */
function forceClearCache() {
  try {
    const piNumber = '13'; // Update this to match your PI
    
    console.log('Clearing cache for PI', piNumber);
    
    if (typeof CacheManager !== 'undefined' && CacheManager.isEnabled()) {
      CacheManager.clearPI(piNumber);
      console.log('✅ Cache cleared successfully');
      console.log('Next trigger run will fetch fresh data regardless of schedule');
    } else {
      console.log('⚠️ Cache manager not available or disabled');
    }
  } catch (error) {
    console.error('❌ Error clearing cache:', error);
  }
}

/**
 * View all active triggers
 */
function listActiveTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  
  console.log('========================================');
  console.log(`ACTIVE TRIGGERS (${triggers.length} total)`);
  console.log('========================================');
  
  if (triggers.length === 0) {
    console.log('No triggers found!');
    console.log('You need to create a trigger in the Apps Script UI:');
    console.log('1. Click the clock icon (Triggers)');
    console.log('2. Add Trigger');
    console.log('3. Function: hourlyUpdate');
    console.log('4. Event source: Time-driven');
    console.log('5. Type: Minute timer');
    console.log('6. Interval: Every 15 minutes');
  } else {
    triggers.forEach((trigger, index) => {
      console.log(`\nTrigger ${index + 1}:`);
      console.log('  Function:', trigger.getHandlerFunction());
      console.log('  Event Type:', trigger.getEventType());
      
      try {
        const triggerSource = trigger.getTriggerSource();
        console.log('  Source:', triggerSource);
      } catch (e) {
        // Some triggers don't have a source
      }
      
      console.log('  Trigger ID:', trigger.getUniqueId());
    });
  }
  
  console.log('========================================');
}

/**
 * Delete all triggers (use with caution!)
 */
function deleteAllTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  
  console.log(`Found ${triggers.length} triggers to delete`);
  
  triggers.forEach(trigger => {
    const functionName = trigger.getHandlerFunction();
    console.log(`Deleting trigger for function: ${functionName}`);
    ScriptApp.deleteTrigger(trigger);
  });
  
  console.log('✅ All triggers deleted');
  console.log('You will need to recreate your 15-minute trigger manually');
}

/**
 * ========================================
 * RATE LIMIT MONITORING
 * ========================================
 */

/**
 * Estimate API calls per hour with current setup
 */
function estimateAPIUsage() {
  console.log('========================================');
  console.log('API USAGE ESTIMATE (Option A)');
  console.log('========================================');
  console.log('Trigger frequency: Every 15 minutes');
  console.log('Runs per hour: 4');
  console.log('');
  console.log('With Option A caching:');
  console.log('  - JIRA API calls: 2 per hour (at :00 and :30)');
  console.log('  - Cache hits: 2 per hour (at :15 and :45)');
  console.log('');
  console.log('Daily totals:');
  console.log('  - Total runs: 96 per day');
  console.log('  - JIRA fetches: 48 per day');
  console.log('  - Cache hits: 48 per day');
  console.log('');
  console.log('This is well within typical JIRA API rate limits!');
  console.log('========================================');
}