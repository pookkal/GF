/**
* ==============================================================================
* STABLE_MASTER_ALL_CLEAN_v3.1_KIRO_OPTIMIZED
* ==============================================================================
*/
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📈 Institutional Terminal')
    .addItem('🚀 1- FETCH DATA', 'FlushDataSheetAndBuild')
    .addItem('🚀 2. REBUILD ALL SHEETS', 'FlushAllSheetsAndBuild')
    .addSeparator()
    .addItem('3. Build Calculations', 'generateCalculationsSheet')
    .addItem('4. Build Dashboard ', 'generateDashboardSheet')
    .addItem('5. Build Mobile Dashboard ', 'generateMobileReport')
    .addItem('🎯 6. Build BUY CATEGORIES', 'generateBuyCategoriesSheet')
    .addSeparator()
    .addItem('🤖 Generate  Narratives', 'runMasterAnalysis')
    .addSeparator()
    .addItem('📖 Build Reference Guide', 'generateReferenceSheet')
    .addSeparator()
    .addItem('🔔 Start Market Monitor', 'startMarketMonitor')
    .addItem('🔕 Stop Monitor', 'stopMarketMonitor')
    .addItem('📩 Test Alert Now', 'checkSignalsAndSendAlerts')
    .addToUi();
}

function onEdit(e) {
  const range = e.range;
  const sheet = range.getSheet();
  const a1 = range.getA1Notation();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // ------------------------------------------------------------
  // INPUT filters -> refresh dashboard
  // ------------------------------------------------------------
  if (sheet.getName() === "INPUT") {
    // Dashboard refresh triggers (B1 or C1)
    if (a1 === "B1" || a1 === "C1") {
      try {
        ss.toast("Dashboard refreshing...", "⚙️ REFRESH", 6);
        generateDashboardSheet();
        SpreadsheetApp.flush();
      } catch (err) {
        ss.toast("Dashboard filter refresh error: " + err.toString(), "⚠️ FAIL", 6);
      }
      return;
    }

    // Data refresh trigger (E1)
    if (a1 === "E1") {
      try {
        ss.toast("Data refreshing...", "⚙️ REFRESH", 6);
        generateDataSheet();
        SpreadsheetApp.flush();
      } catch (err) {
        ss.toast("Data refresh error: " + err.toString(), "⚠️ FAIL", 6);
      }
      return;
    }

    // Calculations refresh trigger (G1) to swithc Invest / Trade mode
    if (a1 === "G1") {
      try {
        ss.toast("Calculations refreshing...", "⚙️ REFRESH", 6);
        generateCalculationsSheet();
        SpreadsheetApp.flush();
      } catch (err) {
        ss.toast("Calculations refresh error: " + err.toString(), "⚠️ FAIL", 6);
      }
      return;
    }
  }

  // ------------------------------------------------------------
  // DASHBOARD update controls:
  // - B1 = Update CALCULATIONS + DASHBOARD
  // - D1 = Update DASHBOARD only
  // ------------------------------------------------------------
  if (sheet.getName() === "DASHBOARD" && (a1 === "B1" || a1 === "D1") && e.value === "TRUE") {
    ss.toast("Refreshing Dashboard...", "⚙️ TERMINAL", 3);
    try {
      if (a1 === "B1") {
        // Full refresh
        generateCalculationsSheet();
      }
      // Dashboard refresh
      generateDashboardSheet();
      ss.toast("Dashboard Synchronized.", "✅ DONE", 2);
    } catch (err) {
      ss.toast("Error: " + err.toString(), "⚠️ FAIL", 6);
    } finally {
      // reset checkbox
      sheet.getRange(a1).setValue(false);
    }
    return;
  }

  // REPORT sheet controls - delegated to generateMobileDashboard.js
  if (sheet.getName() === "REPORT") {
    try {
      if (typeof handleReportSheetEdit === "function") {
        handleReportSheetEdit(e);
      }
    } catch (err) {
      ss.toast("REPORT sheet error: " + err.toString(), "⚠️ FAIL", 6);
    }
    return;
  }

  if (sheet.getName() === "CHART") {
    const watchList = ["A1", "B2", "B3", "B4", "B6"];

    // This triggers if B1-B6 are edited OR any cell in Row 1 (Cols 1-4)
    if (watchList.indexOf(a1) !== -1 || (range.getRow() === 1 && range.getColumn() <= 4)) {
      try {
        ss.toast("🔄 Refreshing Chart & Analysis...", "WORKING", 2);
        if (typeof updateDynamicChart === "function")
          updateDynamicChart();
      } catch (err) {
        ss.toast("Refresh error: " + err.toString(), "⚠️ FAIL", 6);
      }
      return; // Exit after processing CHART
    }
  }
}

//Not used now, This is triggred when CAHR A1 is changed , to call runMasterAnalysis(). This needs setup in triggeres in appscript
function onEditInstall(e) {
  if (!e || !e.range) return;

  const range = e.range;
  const sheet = range.getSheet();

  // Trigger ONLY when CHART!A1 is edited
  //if (sheet.getName() === "CHART" && range.getA1Notation() === "A1") {
  //runMasterAnalysis();
  //}
}

function FlushAllSheetsAndBuild() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetsToDelete = ["CALCULATIONS", "DASHBOARD", "REPORT"];
  const ui = SpreadsheetApp.getUi();

  if (ui.alert('🚨 Full Rebuild', 'Rebuild the sheets?', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;

  sheetsToDelete.forEach(name => {
    let sh = ss.getSheetByName(name);
    if (sh) ss.deleteSheet(sh);
  });

  ui.showModelessDialog(HtmlService.createHtmlOutput("<b>1/3:</b> Integrating Indicators..."), "Status");
  generateCalculationsSheet();
  SpreadsheetApp.flush();

  ui.showModelessDialog(HtmlService.createHtmlOutput("<b>2/3:</b> Building Dashboard..."), "Status");
  generateDashboardSheet();
  SpreadsheetApp.flush();

  ui.showModelessDialog(HtmlService.createHtmlOutput("<b>3/3:</b> Building Mobile Report..."), "Status");
  generateMobileReport();

  ui.showModelessDialog(HtmlService.createHtmlOutput("<b>3/3:</b>✅ Rebuild Complete..."), "Status");
}

function FlushDataSheetAndBuild() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetsToDelete = ["DATA"];
  const ui = SpreadsheetApp.getUi();

  //if (ui.alert('🚨 Full Rebuild', 'Rebuild Data?', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;

  sheetsToDelete.forEach(name => {
    let sh = ss.getSheetByName(name);
    if (sh) ss.deleteSheet(sh);
  });

  ui.showModelessDialog(HtmlService.createHtmlOutput("<b>1/4:</b> Syncing Global Data..."), "Status");
  generateDataSheet();
  SpreadsheetApp.flush();

  ui.alert('✅ Rebuild Complete', 'Data  rerestored.', ui.ButtonSet.OK);
}