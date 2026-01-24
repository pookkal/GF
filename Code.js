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
    .addItem('⚙️ Setup Monitor Config', 'setupMonitorConfiguration')
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

  }

  // ------------------------------------------------------------
  // DASHBOARD control row handlers
  // ------------------------------------------------------------
  if (sheet.getName() === "DASHBOARD") {
    // Country filters (B1, D1)
    if (a1 === "B1" || a1 === "D1") {
      try {
        updateCountryFilter();
      } catch (err) {
        ss.toast("Country filter error: " + err.toString(), "⚠️ FAIL", 3);
      }
      return;
    }

    // Category filter (F1)
    if (a1 === "F1") {
      try {
        updateCategoryFilter();
      } catch (err) {
        ss.toast("Category filter error: " + err.toString(), "⚠️ FAIL", 3);
      }
      return;
    }

    // Mode toggle (H1)
    if (a1 === "H1") {
      try {
        syncModeToggle("DASHBOARD", "H1");
      } catch (err) {
        ss.toast("Mode toggle error: " + err.toString(), "⚠️ FAIL", 3);
      }
      return;
    }

    // Dashboard refresh (J1)
    if (a1 === "J1" && e.value === "TRUE") {
      try {
        refreshDashboardDataFromCheckbox();
      } catch (err) {
        ss.toast("Dashboard refresh error: " + err.toString(), "⚠️ FAIL", 3);
      }
      return;
    }

    // Calculations refresh (L1)
    if (a1 === "L1" && e.value === "TRUE") {
      try {
        refreshCalculations();
      } catch (err) {
        ss.toast("Calculations refresh error: " + err.toString(), "⚠️ FAIL", 3);
      }
      return;
    }

    // Data rebuild (N1)
    if (a1 === "N1" && e.value === "TRUE") {
      try {
        rebuildDataSheet();
      } catch (err) {
        ss.toast("Data rebuild error: " + err.toString(), "⚠️ FAIL", 3);
      }
      return;
    }

    // Sort column change (B2)
    if (a1 === "B2") {
      try {
        onSortColumnChange();
      } catch (err) {
        ss.toast("Sort error: " + err.toString(), "⚠️ FAIL", 3);
      }
      return;
    }
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