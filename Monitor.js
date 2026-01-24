/**
 * ==============================================================================
 * MARKET MONITOR - Daily Change Detection & Email Alerts
 * ==============================================================================
 * Monitors DECISION, SIGNAL, and PATTERN changes daily at 3 AM CET
 * Sends email alerts to the default Gmail account associated with the sheet
 * ==============================================================================
 */

// ============================================================================
// MENU FUNCTIONS (Called from Code.js onOpen menu)
// ============================================================================

/**
 * Start the daily market monitor (creates time-based trigger for 3 AM CET)
 */
function startMarketMonitor() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  
  try {
    // Delete any existing monitor triggers first
    deleteMonitorTriggers_();
    
    // Create new trigger for 3 AM CET (2 AM UTC in winter, 1 AM UTC in summer)
    // Using 2 AM UTC as base (adjust for your timezone)
    ScriptApp.newTrigger('checkSignalsAndSendAlerts')
      .timeBased()
      .atHour(2) // 2 AM UTC = 3 AM CET (winter) / 4 AM CEST (summer)
      .everyDays(1)
      .create();
    
    ss.toast('✅ Market Monitor started - Daily alerts at 3 AM CET', 'Monitor Active', 5);
    ui.alert('🔔 Monitor Started', 
             'Daily monitoring activated!\n\n' +
             '• Checks: DECISION, SIGNAL, PATTERN changes\n' +
             '• Schedule: 3 AM CET daily\n' +
             '• Email: Sent to sheet owner\n\n' +
             'Use "🔕 Stop Monitor" to disable.', 
             ui.ButtonSet.OK);
    
  } catch (error) {
    Logger.log('Error starting monitor: ' + error.toString());
    ss.toast('❌ Failed to start monitor: ' + error.message, 'Error', 5);
  }
}

/**
 * Stop the daily market monitor (removes all triggers)
 */
function stopMarketMonitor() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  
  try {
    const deleted = deleteMonitorTriggers_();
    
    if (deleted > 0) {
      ss.toast('🔕 Market Monitor stopped', 'Monitor Inactive', 3);
      ui.alert('🔕 Monitor Stopped', 
               'Daily monitoring deactivated.\n\n' +
               deleted + ' trigger(s) removed.\n\n' +
               'Use "🔔 Start Market Monitor" to re-enable.', 
               ui.ButtonSet.OK);
    } else {
      ss.toast('ℹ️ No active monitor found', 'Info', 3);
      ui.alert('ℹ️ No Active Monitor', 
               'No monitoring triggers were found.\n\n' +
               'The monitor may already be stopped.', 
               ui.ButtonSet.OK);
    }
    
  } catch (error) {
    Logger.log('Error stopping monitor: ' + error.toString());
    ss.toast('❌ Failed to stop monitor: ' + error.message, 'Error', 5);
  }
}

/**
 * Test alert function - sends immediate alert with current changes
 * Called from menu: "📩 Test Alert Now"
 */
function checkSignalsAndSendAlerts() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  try {
    ss.toast('🔍 Checking for changes...', 'Scanning', 3);
    
    // Run the main monitoring function
    const result = monitorDailyChanges();
    
    if (result.success) {
      ss.toast('✅ Alert sent: ' + result.changeCount + ' changes detected', 'Complete', 5);
    } else {
      ss.toast('⚠️ ' + result.message, 'Warning', 5);
    }
    
  } catch (error) {
    Logger.log('Error in checkSignalsAndSendAlerts: ' + error.toString());
    ss.toast('❌ Alert failed: ' + error.message, 'Error', 5);
  }
}

// ============================================================================
// CORE MONITORING FUNCTIONS
// ============================================================================

/**
 * Main monitoring function - compares current vs previous snapshot
 * Sends email if changes detected above threshold
 * @returns {Object} Result object with success status and message
 */
function monitorDailyChanges() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const startTime = new Date();
  
  try {
    Logger.log('=== MARKET MONITOR START ===');
    Logger.log('Timestamp: ' + startTime.toISOString());
    
    // Get alert threshold from INPUT sheet (K1/L1)
    const threshold = getAlertThreshold_();
    Logger.log('Alert threshold: ' + threshold);
    
    // Capture current snapshot from CALCULATIONS sheet
    const currentSnapshot = captureCurrentSnapshot_();
    if (!currentSnapshot || currentSnapshot.length === 0) {
      Logger.log('No data in CALCULATIONS sheet');
      return {success: false, message: 'No data to monitor', changeCount: 0};
    }
    Logger.log('Current snapshot: ' + currentSnapshot.length + ' tickers');
    
    // Load previous snapshot from MONITOR_HISTORY
    const previousSnapshot = loadPreviousSnapshot_();
    Logger.log('Previous snapshot: ' + (previousSnapshot ? previousSnapshot.length : 0) + ' tickers');
    
    // Detect changes between snapshots
    const changes = detectChanges_(currentSnapshot, previousSnapshot, threshold);
    Logger.log('Changes detected: ' + changes.length);
    
    // Send email alert if changes found
    let emailSent = false;
    if (changes.length > 0) {
      emailSent = sendEmailAlert_(changes, currentSnapshot.length, threshold);
      Logger.log('Email sent: ' + emailSent);
    }
    
    // Save current snapshot as new baseline
    saveSnapshot_(currentSnapshot);
    Logger.log('Snapshot saved to MONITOR_HISTORY');
    
    const elapsed = ((new Date() - startTime) / 1000).toFixed(2);
    Logger.log('=== MONITOR COMPLETE in ' + elapsed + 's ===');
    
    return {
      success: true,
      message: changes.length > 0 ? 'Changes detected and alert sent' : 'No changes detected',
      changeCount: changes.length,
      emailSent: emailSent
    };
    
  } catch (error) {
    Logger.log('ERROR in monitorDailyChanges: ' + error.toString());
    Logger.log('Stack: ' + error.stack);
    return {success: false, message: error.message, changeCount: 0};
  }
}

/**
 * Capture current DECISION, SIGNAL, PATTERN values from CALCULATIONS sheet
 * @returns {Array} Array of objects with ticker and current values
 */
function captureCurrentSnapshot_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const calcSheet = ss.getSheetByName('CALCULATIONS');
  
  if (!calcSheet) {
    throw new Error('CALCULATIONS sheet not found');
  }
  
  const lastRow = calcSheet.getLastRow();
  if (lastRow < 3) {
    return [];
  }
  
  // Read columns: A (Ticker), C (DECISION), D (SIGNAL), E (PATTERNS)
  const data = calcSheet.getRange(3, 1, lastRow - 2, 5).getValues();
  
  const snapshot = [];
  for (let i = 0; i < data.length; i++) {
    const ticker = String(data[i][0] || '').trim().toUpperCase();
    if (!ticker) continue;
    
    const decision = String(data[i][2] || '').trim();
    const signal = String(data[i][3] || '').trim();
    const pattern = String(data[i][4] || '').trim();
    
    // Skip if all values are empty or "LOADING"
    if (!decision && !signal && !pattern) continue;
    if (decision === 'LOADING' && !signal && !pattern) continue;
    
    snapshot.push({
      ticker: ticker,
      decision: decision,
      signal: signal,
      pattern: pattern
    });
  }
  
  return snapshot;
}

/**
 * Load previous snapshot from MONITOR_HISTORY sheet
 * @returns {Array|null} Previous snapshot or null if not found
 */
function loadPreviousSnapshot_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let historySheet = ss.getSheetByName('MONITOR_HISTORY');
  
  if (!historySheet) {
    // Create sheet if it doesn't exist
    historySheet = ss.insertSheet('MONITOR_HISTORY');
    historySheet.hideSheet();
    
    // Set up headers
    historySheet.getRange('A1:E1')
      .setValues([['Timestamp', 'Ticker', 'Decision', 'Signal', 'Pattern']])
      .setFontWeight('bold')
      .setBackground('#1565C0')
      .setFontColor('white');
    
    Logger.log('Created new MONITOR_HISTORY sheet');
    return null;
  }
  
  const lastRow = historySheet.getLastRow();
  if (lastRow < 2) {
    return null;
  }
  
  // Read all historical data
  const data = historySheet.getRange(2, 1, lastRow - 1, 5).getValues();
  
  // Find the most recent timestamp
  let latestTimestamp = null;
  let latestIndex = -1;
  
  for (let i = 0; i < data.length; i++) {
    const timestamp = data[i][0];
    if (timestamp && (!latestTimestamp || timestamp > latestTimestamp)) {
      latestTimestamp = timestamp;
      latestIndex = i;
    }
  }
  
  if (latestIndex === -1) {
    return null;
  }
  
  // Collect all rows with the latest timestamp
  const snapshot = [];
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] && data[i][0].getTime() === latestTimestamp.getTime()) {
      const ticker = String(data[i][1] || '').trim().toUpperCase();
      if (ticker) {
        snapshot.push({
          ticker: ticker,
          decision: String(data[i][2] || '').trim(),
          signal: String(data[i][3] || '').trim(),
          pattern: String(data[i][4] || '').trim()
        });
      }
    }
  }
  
  Logger.log('Loaded previous snapshot from: ' + latestTimestamp.toISOString());
  return snapshot;
}

/**
 * Detect changes between current and previous snapshots
 * @param {Array} current - Current snapshot
 * @param {Array} previous - Previous snapshot
 * @param {string} threshold - Alert threshold (ALL, CRITICAL, CUSTOM)
 * @returns {Array} Array of change objects
 */
function detectChanges_(current, previous, threshold) {
  if (!previous || previous.length === 0) {
    Logger.log('No previous snapshot - treating all as new');
    return []; // Don't alert on first run
  }
  
  // Create lookup map for previous values
  const prevMap = {};
  previous.forEach(item => {
    prevMap[item.ticker] = item;
  });
  
  const changes = [];
  
  current.forEach(curr => {
    const prev = prevMap[curr.ticker];
    
    if (!prev) {
      // New ticker - only alert if threshold is ALL
      if (threshold === 'ALL') {
        changes.push({
          ticker: curr.ticker,
          type: 'NEW',
          field: 'ALL',
          oldValue: '',
          newValue: 'New ticker added',
          priority: 'LOW'
        });
      }
      return;
    }
    
    // Check DECISION changes
    if (curr.decision !== prev.decision && 
        curr.decision !== 'LOADING' && 
        prev.decision !== 'LOADING') {
      
      const priority = getChangePriority_('DECISION', prev.decision, curr.decision);
      
      if (shouldAlert_(threshold, priority)) {
        changes.push({
          ticker: curr.ticker,
          type: 'DECISION',
          field: 'DECISION',
          oldValue: prev.decision,
          newValue: curr.decision,
          priority: priority
        });
      }
    }
    
    // Check SIGNAL changes
    if (curr.signal !== prev.signal && 
        curr.signal !== 'LOADING' && 
        prev.signal !== 'LOADING') {
      
      const priority = getChangePriority_('SIGNAL', prev.signal, curr.signal);
      
      if (shouldAlert_(threshold, priority)) {
        changes.push({
          ticker: curr.ticker,
          type: 'SIGNAL',
          field: 'SIGNAL',
          oldValue: prev.signal,
          newValue: curr.signal,
          priority: priority
        });
      }
    }
    
    // Check PATTERN changes
    if (curr.pattern !== prev.pattern && curr.pattern && prev.pattern) {
      const priority = 'MEDIUM'; // Patterns are always medium priority
      
      if (shouldAlert_(threshold, priority)) {
        changes.push({
          ticker: curr.ticker,
          type: 'PATTERN',
          field: 'PATTERN',
          oldValue: prev.pattern,
          newValue: curr.pattern,
          priority: priority
        });
      }
    }
  });
  
  // Sort by priority: CRITICAL > HIGH > MEDIUM > LOW
  const priorityOrder = {CRITICAL: 0, HIGH: 1, MEDIUM: 2, LOW: 3};
  changes.sort((a, b) => priorityOrder[a.priority] - priorityOrder[b.priority]);
  
  return changes;
}

/**
 * Determine change priority based on field and values
 * @param {string} field - Field name (DECISION, SIGNAL, PATTERN)
 * @param {string} oldVal - Previous value
 * @param {string} newVal - New value
 * @returns {string} Priority level (CRITICAL, HIGH, MEDIUM, LOW)
 */
function getChangePriority_(field, oldVal, newVal) {
  const old = oldVal.toUpperCase();
  const newV = newVal.toUpperCase();
  
  if (field === 'DECISION') {
    // CRITICAL: Major reversals
    if ((old.includes('BUY') && newV.includes('SELL')) ||
        (old.includes('SELL') && newV.includes('BUY')) ||
        (old.includes('STRONG') && newV.includes('STOP')) ||
        (newV.includes('STOP OUT'))) {
      return 'CRITICAL';
    }
    
    // HIGH: Significant changes
    if ((old.includes('HOLD') && newV.includes('BUY')) ||
        (old.includes('BUY') && newV.includes('HOLD')) ||
        (newV.includes('STRONG BUY')) ||
        (old.includes('STRONG') && !newV.includes('STRONG'))) {
      return 'HIGH';
    }
    
    // MEDIUM: Moderate changes
    return 'MEDIUM';
  }
  
  if (field === 'SIGNAL') {
    // CRITICAL: Risk signals
    if (newV.includes('RISK OFF') || newV.includes('STOP OUT') || newV.includes('TRIM')) {
      return 'CRITICAL';
    }
    
    // HIGH: Strong buy/sell signals
    if (newV.includes('STRONG BUY') || newV.includes('STRONG SELL') || 
        newV.includes('ACCUMULATE') || newV.includes('BREAKOUT')) {
      return 'HIGH';
    }
    
    // MEDIUM: Other signals
    return 'MEDIUM';
  }
  
  // PATTERN changes are always MEDIUM
  return 'MEDIUM';
}

/**
 * Check if change should trigger alert based on threshold
 * @param {string} threshold - Alert threshold setting
 * @param {string} priority - Change priority
 * @returns {boolean} True if should alert
 */
function shouldAlert_(threshold, priority) {
  if (threshold === 'All' || threshold === 'ALL') return true;
  if (threshold === 'CRITICAL') return priority === 'CRITICAL';
  if (threshold === 'HIGH') return priority === 'CRITICAL' || priority === 'HIGH';
  return true; // Default to All
}

/**
 * Send email alert with detected changes
 * @param {Array} changes - Array of change objects
 * @param {number} totalTickers - Total number of tickers monitored
 * @param {string} threshold - Alert threshold used
 * @returns {boolean} True if email sent successfully
 */
function sendEmailAlert_(changes, totalTickers, threshold) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const email = Session.getActiveUser().getEmail();
    
    if (!email) {
      Logger.log('No email address found');
      return false;
    }
    
    const subject = '📊 Market Monitor Alert - ' + changes.length + ' Change' + 
                    (changes.length !== 1 ? 's' : '') + ' Detected';
    
    const htmlBody = buildEmailHTML_(changes, totalTickers, threshold);
    
    MailApp.sendEmail({
      to: email,
      subject: subject,
      htmlBody: htmlBody
    });
    
    Logger.log('Email sent to: ' + email);
    return true;
    
  } catch (error) {
    Logger.log('Error sending email: ' + error.toString());
    return false;
  }
}

/**
 * Build HTML email body with formatted changes
 * @param {Array} changes - Array of change objects
 * @param {number} totalTickers - Total tickers monitored
 * @param {string} threshold - Alert threshold
 * @returns {string} HTML email body
 */
function buildEmailHTML_(changes, totalTickers, threshold) {
  const timestamp = Utilities.formatDate(new Date(), 
                                         SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), 
                                         'yyyy-MM-dd HH:mm:ss z');
  
  // Count by priority
  const criticalCount = changes.filter(c => c.priority === 'CRITICAL').length;
  const highCount = changes.filter(c => c.priority === 'HIGH').length;
  const mediumCount = changes.filter(c => c.priority === 'MEDIUM').length;
  
  // Group changes by type
  const decisionChanges = changes.filter(c => c.type === 'DECISION');
  const signalChanges = changes.filter(c => c.type === 'SIGNAL');
  const patternChanges = changes.filter(c => c.type === 'PATTERN');
  
  let html = `
<!DOCTYPE html>
<html>
<head>
  <style>
    body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Arial, sans-serif; line-height: 1.6; color: #333; max-width: 800px; margin: 0 auto; padding: 20px; }
    .header { background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; padding: 30px; border-radius: 10px; margin-bottom: 30px; }
    .header h1 { margin: 0; font-size: 28px; }
    .header p { margin: 10px 0 0 0; opacity: 0.9; }
    .summary { background: #f8f9fa; border-left: 4px solid #667eea; padding: 20px; margin-bottom: 30px; border-radius: 5px; }
    .summary-grid { display: grid; grid-template-columns: repeat(2, 1fr); gap: 15px; margin-top: 15px; }
    .summary-item { background: white; padding: 15px; border-radius: 5px; }
    .summary-item strong { display: block; font-size: 24px; color: #667eea; }
    .section { margin-bottom: 30px; }
    .section-title { font-size: 20px; font-weight: bold; color: #667eea; margin-bottom: 15px; padding-bottom: 10px; border-bottom: 2px solid #e9ecef; }
    .change-item { background: white; border: 1px solid #e9ecef; border-radius: 8px; padding: 15px; margin-bottom: 10px; }
    .change-header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 10px; }
    .ticker { font-size: 18px; font-weight: bold; color: #2c3e50; }
    .priority { padding: 4px 12px; border-radius: 12px; font-size: 12px; font-weight: bold; text-transform: uppercase; }
    .priority-critical { background: #fee; color: #c00; }
    .priority-high { background: #fff3cd; color: #856404; }
    .priority-medium { background: #d1ecf1; color: #0c5460; }
    .priority-low { background: #e2e3e5; color: #383d41; }
    .change-detail { display: flex; align-items: center; gap: 10px; font-size: 14px; }
    .arrow { color: #667eea; font-weight: bold; }
    .old-value { color: #6c757d; text-decoration: line-through; }
    .new-value { color: #28a745; font-weight: bold; }
    .footer { text-align: center; color: #6c757d; font-size: 12px; margin-top: 40px; padding-top: 20px; border-top: 1px solid #e9ecef; }
  </style>
</head>
<body>
  <div class="header">
    <h1>📊 Market Monitor Alert</h1>
    <p>${timestamp}</p>
  </div>
  
  <div class="summary">
    <strong>Summary</strong>
    <div class="summary-grid">
      <div class="summary-item">
        <strong>${totalTickers}</strong>
        <span>Tickers Monitored</span>
      </div>
      <div class="summary-item">
        <strong>${changes.length}</strong>
        <span>Changes Detected</span>
      </div>
      <div class="summary-item">
        <strong>${criticalCount}</strong>
        <span>Critical Alerts</span>
      </div>
      <div class="summary-item">
        <strong>${threshold}</strong>
        <span>Alert Threshold</span>
      </div>
    </div>
  </div>
`;
  
  // DECISION Changes
  if (decisionChanges.length > 0) {
    html += `
  <div class="section">
    <div class="section-title">🎯 DECISION Changes (${decisionChanges.length})</div>
`;
    decisionChanges.forEach(change => {
      html += formatChangeItem_(change);
    });
    html += `  </div>\n`;
  }
  
  // SIGNAL Changes
  if (signalChanges.length > 0) {
    html += `
  <div class="section">
    <div class="section-title">📡 SIGNAL Changes (${signalChanges.length})</div>
`;
    signalChanges.forEach(change => {
      html += formatChangeItem_(change);
    });
    html += `  </div>\n`;
  }
  
  // PATTERN Changes
  if (patternChanges.length > 0) {
    html += `
  <div class="section">
    <div class="section-title">📈 PATTERN Changes (${patternChanges.length})</div>
`;
    patternChanges.forEach(change => {
      html += formatChangeItem_(change);
    });
    html += `  </div>\n`;
  }
  
  html += `
  <div class="footer">
    <p>This is an automated alert from your Market Monitor system.</p>
    <p>To manage alerts, use the "📈 Institutional Terminal" menu in your spreadsheet.</p>
  </div>
</body>
</html>
`;
  
  return html;
}

/**
 * Format individual change item for email
 * @param {Object} change - Change object
 * @returns {string} HTML for change item
 */
function formatChangeItem_(change) {
  const priorityClass = 'priority-' + change.priority.toLowerCase();
  
  return `
    <div class="change-item">
      <div class="change-header">
        <span class="ticker">${change.ticker}</span>
        <span class="priority ${priorityClass}">${change.priority}</span>
      </div>
      <div class="change-detail">
        <span class="old-value">${change.oldValue || '—'}</span>
        <span class="arrow">→</span>
        <span class="new-value">${change.newValue}</span>
      </div>
    </div>
`;
}

/**
 * Save current snapshot to MONITOR_HISTORY sheet
 * @param {Array} snapshot - Current snapshot to save
 */
function saveSnapshot_(snapshot) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let historySheet = ss.getSheetByName('MONITOR_HISTORY');
  
  if (!historySheet) {
    historySheet = ss.insertSheet('MONITOR_HISTORY');
    historySheet.hideSheet();
    
    // Set up headers
    historySheet.getRange('A1:E1')
      .setValues([['Timestamp', 'Ticker', 'Decision', 'Signal', 'Pattern']])
      .setFontWeight('bold')
      .setBackground('#1565C0')
      .setFontColor('white');
  }
  
  const timestamp = new Date();
  const data = snapshot.map(item => [
    timestamp,
    item.ticker,
    item.decision,
    item.signal,
    item.pattern
  ]);
  
  // Append new snapshot
  if (data.length > 0) {
    const lastRow = historySheet.getLastRow();
    historySheet.getRange(lastRow + 1, 1, data.length, 5).setValues(data);
  }
  
  // Clean up old data (keep last 30 days)
  cleanupOldHistory_(historySheet);
}

/**
 * Remove history older than 30 days
 * @param {Sheet} historySheet - MONITOR_HISTORY sheet
 */
function cleanupOldHistory_(historySheet) {
  const lastRow = historySheet.getLastRow();
  if (lastRow < 2) return;
  
  const cutoffDate = new Date();
  cutoffDate.setDate(cutoffDate.getDate() - 30);
  
  const timestamps = historySheet.getRange(2, 1, lastRow - 1, 1).getValues();
  
  // Find rows to delete (from bottom to avoid index shifting)
  const rowsToDelete = [];
  for (let i = timestamps.length - 1; i >= 0; i--) {
    if (timestamps[i][0] && timestamps[i][0] < cutoffDate) {
      rowsToDelete.push(i + 2); // +2 because array is 0-indexed and data starts at row 2
    }
  }
  
  // Delete old rows
  rowsToDelete.forEach(row => {
    historySheet.deleteRow(row);
  });
  
  if (rowsToDelete.length > 0) {
    Logger.log('Cleaned up ' + rowsToDelete.length + ' old history rows');
  }
}

/**
 * Get alert threshold from DASHBOARD sheet (P1)
 * @returns {string} Threshold value (All, CRITICAL, HIGH)
 */
function getAlertThreshold_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dashboard = ss.getSheetByName('DASHBOARD');
  
  if (!dashboard) {
    return 'All'; // Default
  }
  
  const threshold = String(dashboard.getRange('P1').getValue() || 'All').trim();
  
  // Validate threshold
  if (['All', 'CRITICAL', 'HIGH'].indexOf(threshold) === -1) {
    return 'All';
  }
  
  return threshold;
}

/**
 * Delete all monitor-related triggers
 * @returns {number} Number of triggers deleted
 */
function deleteMonitorTriggers_() {
  const triggers = ScriptApp.getProjectTriggers();
  let deleted = 0;
  
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'checkSignalsAndSendAlerts') {
      ScriptApp.deleteTrigger(trigger);
      deleted++;
    }
  });
  
  Logger.log('Deleted ' + deleted + ' monitor trigger(s)');
  return deleted;
}


// ============================================================================
// SETUP FUNCTION - Initialize INPUT sheet with threshold configuration
// ============================================================================

/**
 * Setup alert threshold configuration in INPUT sheet (K1/L1)
 * Call this once to initialize the configuration
 */
function setupMonitorConfiguration() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dashboard = ss.getSheetByName('DASHBOARD');
  
  if (!dashboard) {
    ss.toast('❌ DASHBOARD sheet not found', 'Error', 3);
    return;
  }
  
  try {
    // Alert threshold is now in DASHBOARD P1
    // This is automatically set up by setupControlRowDropdowns()
    
    ss.toast('✅ Alert threshold is configured in DASHBOARD P1', 'Info', 5);
    
    const ui = SpreadsheetApp.getUi();
    ui.alert('ℹ️ Alert Configuration', 
             'Alert threshold is now controlled from the DASHBOARD sheet.\n\n' +
             '• Location: DASHBOARD P1 (ALERT dropdown)\n' +
             '• Options: All, HIGH, CRITICAL\n' +
             '• Default: All\n\n' +
             'The dropdown is automatically configured when you build the dashboard.', 
             ui.ButtonSet.OK);
    
    Logger.log('Monitor configuration info displayed');
    
  } catch (error) {
    Logger.log('Error in setupMonitorConfiguration: ' + error.toString());
    ss.toast('❌ Setup failed: ' + error.message, 'Error', 5);
  }
}
