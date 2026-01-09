/**
* ==============================================================================
* STABLE_MASTER_ALL_CLEAN_v2.1._INVEST
* ==============================================================================
*/

/**
* ------------------------------------------------------------------
*  Open LOGIC ENGINE (INSERT MENU)
* ------------------------------------------------------------------
*/
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📈 Institutional Terminal')
    .addItem('🚀 1- FETCH DATA', 'FlushDataSheetAndBuild')
    .addItem('🚀 2. REBUILD ALL SHEETS', 'FlushAllSheetsAndBuild')
    .addSeparator()
    .addItem('3. Build Calculations', 'generateCalculationsSheet')
    .addItem('4. Refresh Dashboard ', 'generateDashboardSheet')
    .addItem('4. Refresh Mobile Dashbaord ', 'setupFormulaBasedReport') //generateMobileReport
    .addItem('5. Setup Chart', 'setupChartSheet')
    .addSeparator()
    .addItem('🤖 Generate  Narratives', 'runMasterAnalysis')
    .addSeparator()
    .addItem('📖 Open Reference Guide', 'generateReferenceSheet')
    .addSeparator()
    .addItem('🔔 Start Market Monitor', 'startMarketMonitor')
    .addItem('🔕 Stop Monitor', 'stopMarketMonitor')
    .addItem('📩 Test Alert Now', 'checkSignalsAndSendAlerts')
    .addToUi();
}

// ------------------------------------------------------------
// UPDATED onEdit(e) — watches the changes to update shets
// ------------------------------------------------------------

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

    // Calculations refresh trigger (E2)
    if (a1 === "E2") {
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

  // REPORT sheet controls - consolidated block
  if (sheet.getName() === "REPORT") {
    const row = range.getRow();
    const col = range.getColumn();
    
    // Handle chart controls: checkbox changes (row 2, columns E-M: 5-13), ticker change (A1), or date/interval change (A2:C2 and C3)
    if ((row === 2 && col >= 5 && col <= 13) || a1 === "A1" || (row === 2 && col >= 1 && col <= 3) || a1 === "C3") {
      try {
        ss.toast("🔄 Updating REPORT Chart...", "WORKING", 2);
        updateReportChart();
      } catch (err) {
        ss.toast("REPORT Chart update error: " + err.toString(), "⚠️ FAIL", 6);
      }
      return;
    }
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

function onEditInstall(e) {
  if (!e || !e.range) return;

  const range = e.range;
  const sheet = range.getSheet();

  // Trigger ONLY when CHART!A1 is edited
  //if (sheet.getName() === "CHART" && range.getA1Notation() === "A1") {
  //runMasterAnalysis();
  //}
}
/**
* ------------------------------------------------------------------
* 1. CORE AUTOMATION
* ------------------------------------------------------------------
*/
function FlushAllSheetsAndBuild() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetsToDelete = ["CALCULATIONS", "DASHBOARD", "CHART", "REPORT"];
  const ui = SpreadsheetApp.getUi();

  if (ui.alert('🚨 Full Rebuild', 'Rebuild the sheets?', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;

  sheetsToDelete.forEach(name => {
    let sh = ss.getSheetByName(name);
    if (sh) ss.deleteSheet(sh);
  });

  ui.showModelessDialog(HtmlService.createHtmlOutput("<b>1/4:</b> Integrating Indicators..."), "Status");
  generateCalculationsSheet();
  SpreadsheetApp.flush();

  ui.showModelessDialog(HtmlService.createHtmlOutput("<b>2/4:</b> Building Dashboard..."), "Status");
  generateDashboardSheet();
  SpreadsheetApp.flush();

  ui.showModelessDialog(HtmlService.createHtmlOutput("<b>3/4:</b> Constructing Report..."), "Status");
  setupFormulaBasedReport();

  ui.showModelessDialog(HtmlService.createHtmlOutput("<b>4/4:</b> Constructing Chart..."), "Status");
  setupChartSheet();

  ui.alert('✅ Rebuild Complete', 'Terminal Online. Data links restored.', ui.ButtonSet.OK);
}

function FlushDataSheetAndBuild() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetsToDelete = ["DATA"];
  const ui = SpreadsheetApp.getUi();

  if (ui.alert('🚨 Full Rebuild', 'Rebuild Data?', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;

  sheetsToDelete.forEach(name => {
    let sh = ss.getSheetByName(name);
    if (sh) ss.deleteSheet(sh);
  });

  ui.showModelessDialog(HtmlService.createHtmlOutput("<b>1/4:</b> Syncing Global Data..."), "Status");
  generateDataSheet();
  SpreadsheetApp.flush();

  ui.alert('✅ Rebuild Complete', 'Data  rerestored.', ui.ButtonSet.OK);
}

/**
* ------------------------------------------------------------------
* 3. DATA ENGINE (FULL FUNCTION — ROW 2 TICKER, ROW 3 ATH/PE/EPS IN A..F)
* ------------------------------------------------------------------
*/
function generateDataSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inputSheet = ss.getSheetByName("INPUT");
  if (!inputSheet) return;

  const tickers = getCleanTickers(inputSheet);
  let dataSheet = ss.getSheetByName("DATA") || ss.insertSheet("DATA");

  // Clear
  dataSheet.clear({ contentsOnly: true });
  dataSheet.clearFormats();

  // Timestamp
  dataSheet.getRange("A1")
    .setValue("Last Update: " + Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "yyyy-MM-dd HH:mm"))
    .setFontWeight("bold")
    .setFontColor("blue");

  if (tickers.length === 0) return;

  const colsPer = 7;
  const totalCols = tickers.length * colsPer;
  
  // MARKET REGIME COLUMNS (add after ticker columns)
  const regimeStartCol = totalCols + 1;
  const regimeColsNeeded = 6; // USA Regime, USA Ratio, USA VIX, India Regime, India Ratio, India VIX
  const finalTotalCols = totalCols + regimeColsNeeded;

  // Ensure enough columns
  if (dataSheet.getMaxColumns() < finalTotalCols) {
    dataSheet.insertColumnsAfter(dataSheet.getMaxColumns(), finalTotalCols - dataSheet.getMaxColumns());
  }

  // ------------------------------------------------------------
  // Row 2: Tickers + Market Regime Headers
  // ------------------------------------------------------------
  const row2 = new Array(finalTotalCols).fill("");
  for (let i = 0; i < tickers.length; i++) {
    row2[i * colsPer] = tickers[i];
  }
  
  // Add market regime headers
  row2[regimeStartCol - 1] = "USA_REGIME";
  row2[regimeStartCol] = "USA_RATIO"; 
  row2[regimeStartCol + 1] = "USA_VIX";
  row2[regimeStartCol + 2] = "INDIA_REGIME";
  row2[regimeStartCol + 3] = "INDIA_RATIO";
  row2[regimeStartCol + 4] = "INDIA_VIX";
  
  dataSheet.getRange(2, 1, 1, finalTotalCols)
    .setValues([row2])
    .setNumberFormat("@")
    .setFontWeight("bold");

  // ------------------------------------------------------------
  // Row 3: Formulas first (ATH / P-E / EPS values + Market Regime)
  // ------------------------------------------------------------
  const row3Formulas = new Array(finalTotalCols).fill("");
  for (let i = 0; i < tickers.length; i++) {
    const t = tickers[i];
    const b = i * colsPer;

    // value cells only
    row3Formulas[b + 1] =
      `=MAX(QUERY(GOOGLEFINANCE("${t}","high","1/1/2000",TODAY()),"SELECT Col2 LABEL Col2 ''"))`;
    row3Formulas[b + 3] =
      `=IFERROR(GOOGLEFINANCE("${t}","pe"),"")`;
    row3Formulas[b + 5] =
      `=IFERROR(GOOGLEFINANCE("${t}","eps"),"")`;
  }
  
  // Add market regime formulas (calculated once for all tickers)
  const SEP = (/^(en|en_)/.test(ss.getSpreadsheetLocale())) ? "," : ";";
  
  // USA Market Regime
  row3Formulas[regimeStartCol - 1] = 
    `=LET(` +
    `spyPrice${SEP}IFERROR(GOOGLEFINANCE("SPY"${SEP}"price")${SEP}0)${SEP}` +
    `spySMA200${SEP}IFERROR(AVERAGE(QUERY(GOOGLEFINANCE("SPY"${SEP}"close"${SEP}TODAY()-250${SEP}TODAY())${SEP}"SELECT Col2 ORDER BY Col1 DESC LIMIT 200"))${SEP}spyPrice)${SEP}` +
    `regimeRatio${SEP}IF(spySMA200>0${SEP}spyPrice/spySMA200${SEP}1)${SEP}` +
    `vixLevel${SEP}IFERROR(GOOGLEFINANCE("INDEXCBOE:VIX"${SEP}"price")${SEP}20)${SEP}` +
    `IFS(` +
    `AND(regimeRatio>=1.05${SEP}vixLevel<=18)${SEP}"STRONG BULL"${SEP}` +
    `AND(regimeRatio>=1.02${SEP}vixLevel<=25)${SEP}"BULL"${SEP}` +
    `AND(regimeRatio>=0.98${SEP}vixLevel<=30)${SEP}"NEUTRAL"${SEP}` +
    `AND(regimeRatio>=0.95${SEP}vixLevel<=35)${SEP}"BEAR"${SEP}` +
    `TRUE${SEP}"STRONG BEAR"` +
    `)` +
    `)`;
    
  row3Formulas[regimeStartCol] = 
    `=IFERROR(GOOGLEFINANCE("SPY"${SEP}"price")/AVERAGE(QUERY(GOOGLEFINANCE("SPY"${SEP}"close"${SEP}TODAY()-250${SEP}TODAY())${SEP}"SELECT Col2 ORDER BY Col1 DESC LIMIT 200"))${SEP}1)`;
    
  row3Formulas[regimeStartCol + 1] = 
    `=IFERROR(GOOGLEFINANCE("INDEXCBOE:VIX"${SEP}"price")${SEP}20)`;
  
  // India Market Regime  
  row3Formulas[regimeStartCol + 2] = 
    `=LET(` +
    `niftyPrice${SEP}IFERROR(GOOGLEFINANCE("INDEXNSE:NIFTY_50"${SEP}"price")${SEP}0)${SEP}` +
    `niftySMA200${SEP}IFERROR(AVERAGE(QUERY(GOOGLEFINANCE("INDEXNSE:NIFTY_50"${SEP}"close"${SEP}TODAY()-250${SEP}TODAY())${SEP}"SELECT Col2 ORDER BY Col1 DESC LIMIT 200"))${SEP}niftyPrice)${SEP}` +
    `regimeRatio${SEP}IF(niftySMA200>0${SEP}niftyPrice/niftySMA200${SEP}1)${SEP}` +
    `vixLevel${SEP}IFERROR(GOOGLEFINANCE("INDEXNSE:INDIAVIX"${SEP}"price")${SEP}20)${SEP}` +
    `IFS(` +
    `AND(regimeRatio>=1.05${SEP}vixLevel<=18)${SEP}"STRONG BULL"${SEP}` +
    `AND(regimeRatio>=1.02${SEP}vixLevel<=25)${SEP}"BULL"${SEP}` +
    `AND(regimeRatio>=0.98${SEP}vixLevel<=30)${SEP}"NEUTRAL"${SEP}` +
    `AND(regimeRatio>=0.95${SEP}vixLevel<=35)${SEP}"BEAR"${SEP}` +
    `TRUE${SEP}"STRONG BEAR"` +
    `)` +
    `)`;
    
  row3Formulas[regimeStartCol + 3] = 
    `=IFERROR(GOOGLEFINANCE("INDEXNSE:NIFTY_50"${SEP}"price")/AVERAGE(QUERY(GOOGLEFINANCE("INDEXNSE:NIFTY_50"${SEP}"close"${SEP}TODAY()-250${SEP}TODAY())${SEP}"SELECT Col2 ORDER BY Col1 DESC LIMIT 200"))${SEP}1)`;
    
  row3Formulas[regimeStartCol + 4] = 
    `=IFERROR(GOOGLEFINANCE("INDEXNSE:INDIAVIX"${SEP}"price")${SEP}20)`;
  
  dataSheet.getRange(3, 1, 1, finalTotalCols).setFormulas([row3Formulas]);

  // Now write labels (cannot be overwritten now)
  for (let i = 0; i < tickers.length; i++) {
    const c = (i * colsPer) + 1; // 1-based
    dataSheet.getRange(3, c).setValue("ATH:");
    dataSheet.getRange(3, c + 2).setValue("P/E:");
    dataSheet.getRange(3, c + 4).setValue("EPS:");
  }

  // ------------------------------------------------------------
  // Row 4: GOOGLEFINANCE(all)
  // ------------------------------------------------------------
  const row4Formulas = new Array(finalTotalCols).fill("");
  for (let i = 0; i < tickers.length; i++) {
    const t = tickers[i];
    row4Formulas[i * colsPer] =
      `=IFERROR(GOOGLEFINANCE("${t}","all",TODAY()-800,TODAY()),"No Data")`;
  }
  dataSheet.getRange(4, 1, 1, finalTotalCols).setFormulas([row4Formulas]);

  // ------------------------------------------------------------
  // Number formats (row 3 values)
  // ------------------------------------------------------------
  for (let i = 0; i < tickers.length; i++) {
    const c = (i * colsPer) + 1; // 1-based
    dataSheet.getRange(3, c + 1).setNumberFormat("#,##0.00"); // ATH value
    dataSheet.getRange(3, c + 3).setNumberFormat("0.00");     // P/E value
    dataSheet.getRange(3, c + 5).setNumberFormat("0.00");     // EPS value
  }
  
  // Market regime formatting
  dataSheet.getRange(3, regimeStartCol, 1, 1).setNumberFormat("@");     // USA_REGIME (text)
  dataSheet.getRange(3, regimeStartCol + 1, 1, 1).setNumberFormat("0.000"); // USA_RATIO (3 decimals)
  dataSheet.getRange(3, regimeStartCol + 2, 1, 1).setNumberFormat("0.0");   // USA_VIX (1 decimal)
  dataSheet.getRange(3, regimeStartCol + 3, 1, 1).setNumberFormat("@");     // INDIA_REGIME (text)
  dataSheet.getRange(3, regimeStartCol + 4, 1, 1).setNumberFormat("0.000"); // INDIA_RATIO (3 decimals)
  dataSheet.getRange(3, regimeStartCol + 5, 1, 1).setNumberFormat("0.0");   // INDIA_VIX (1 decimal)

  // ------------------------------------------------------------
  // Label styling (guaranteed visible)
  // ------------------------------------------------------------
  const LABEL_BG = "#1F2937";
  const LABEL_FG = "#F9FAFB";

  const labelA1s = [];
  for (let i = 0; i < tickers.length; i++) {
    const c = (i * colsPer) + 1; // 1-based
    labelA1s.push(dataSheet.getRange(3, c).getA1Notation());       // ATH label
    labelA1s.push(dataSheet.getRange(3, c + 2).getA1Notation());   // P/E label
    labelA1s.push(dataSheet.getRange(3, c + 4).getA1Notation());   // EPS label
  }

  dataSheet.getRangeList(labelA1s)
    .setBackground(LABEL_BG)
    .setFontColor(LABEL_FG)
    .setFontWeight("bold")
    .setHorizontalAlignment("left");
    
  // Market regime headers styling
  dataSheet.getRange(2, regimeStartCol, 1, regimeColsNeeded)
    .setBackground("#4A148C")  // Purple background for institutional data
    .setFontColor("#FFFFFF")   // White text
    .setFontWeight("bold")
    .setHorizontalAlignment("center");
    
  // Market regime values styling  
  dataSheet.getRange(3, regimeStartCol, 1, regimeColsNeeded)
    .setBackground("#E1BEE7")  // Light purple background
    .setFontColor("#4A148C")   // Dark purple text
    .setFontWeight("bold")
    .setHorizontalAlignment("center");

  // ------------------------------------------------------------
  // Historical formatting (rows 5+)
  // ------------------------------------------------------------
  for (let i = 0; i < tickers.length; i++) {
    const colStart = (i * colsPer) + 1; // 1-based
    dataSheet.getRange(5, colStart, 1000, 1).setNumberFormat("yyyy-mm-dd");
    dataSheet.getRange(5, colStart + 1, 1000, 5).setNumberFormat("#,##0.00");
  }

  SpreadsheetApp.flush();
}


function getCleanTickers(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return [];
  return sheet.getRange(3, 1, lastRow - 2, 1)
    .getValues()
    .flat()
    .filter(t => t && t.toString().trim() !== "")
    .map(t => t.toString().toUpperCase().trim());
}

function columnToLetter(column) {
  let temp, letter = '';
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}

function forceExpandSheet(sheet, targetCols) {
  if (sheet.getMaxColumns() < targetCols) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), targetCols - sheet.getMaxColumns());
  }
}

/**
* ------------------------------------------------------------------
* 4. CALCULATION ENGINE (FULL FUNCTION — UPDATED)
* - Fixes: SELL-side decisions (Take Profit / Reduce)
* - Fixes: Locale separator auto-handled (; vs ,)
* - Formatting: LEFT align + WRAP + row height ~4 lines (72px)
* ------------------------------------------------------------------
* Columns (A..AH):
* A  Ticker
* B  SIGNAL
* C  DECISION
* D  FUNDAMENTAL
* E  Price
* F  Change %
* G  Vol Trend
* H  ATH (TRUE)
* I  ATH Diff %
* J  R:R Quality
* K  Trend Score
* L  Trend State
* M  SMA 20
* N  SMA 50
* O  SMA 200
* P  RSI
* Q  MACD Hist
* R  Divergence
* S  ADX (14)
* T  Stoch %K (14)
* U  Support
* V  Resistance
* W  Target (3:1)
* X  ATR (14)
* Y  Bollinger %B
* Z  POSITION SIZE
* AA TECH NOTES
* AB FUND NOTES
* AC VOL REGIME
* AD ATH ZONE
* AE BBP SIGNAL
* AF PATTERNS
* AG ATR STOP
* AH ATR TARGET
* ------------------------------------------------------------------
*/
function generateCalculationsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getSheetByName("DATA");
  const inputSheet = ss.getSheetByName("INPUT");
  if (!dataSheet || !inputSheet) return;

  const tickers = getCleanTickers(inputSheet);
  let calc = ss.getSheetByName("CALCULATIONS") || ss.insertSheet("CALCULATIONS");

  // Locale separator: US typically ","; EU typically ";"
  const locale = (ss.getSpreadsheetLocale() || "").toLowerCase();
  const SEP = (/^(en|en_)/.test(locale)) ? "," : ";";

  // Persist state map for restoration (no longer using LAST_STATE column)
  const stateMap = {};
  if (calc.getLastRow() >= 3) {
    const existing = calc.getRange(3, 1, calc.getLastRow() - 2, 28).getValues();
    existing.forEach(r => {
      const t = (r[0] || "").toString().trim().toUpperCase();
      // State persistence removed - no longer needed
    });
  }

  calc.clear().clearFormats();

  // Ensure sheet has enough columns for all our data (34 columns total)
  const maxCols = calc.getMaxColumns();
  if (maxCols < 34) {
    calc.insertColumnsAfter(maxCols, 34 - maxCols);
  }

  // ------------------------------------------------------------------
  // ROW 1: GROUP HEADERS (MERGED) + timestamp in AB1
  // ------------------------------------------------------------------
  const syncTime = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "yyyy-MM-dd HH:mm:ss");
  const styleGroup = (a1, label, bg) => {
    calc.getRange(a1).merge()
      .setValue(label)
      .setBackground(bg)
      .setFontColor("white")
      .setFontWeight("bold")
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle");
  };

  styleGroup("A1:A1", "IDENTITY", "#263238"); // A
  styleGroup("B1:D1", "SIGNALING", "#0D47A1"); // B-D
  styleGroup("E1:G1", "PRICE / VOLUME", "#1B5E20"); // E-G
  styleGroup("H1:J1", "PERFORMANCE", "#004D40"); // H-J
  styleGroup("K1:O1", "TREND", "#2E7D32"); // K-O
  styleGroup("P1:T1", "MOMENTUM", "#33691E"); // P-T
  styleGroup("U1:Y1", "LEVELS / RISK", "#B71C1C"); // U-Y
  styleGroup("Z1:Z1", "INSTITUTIONAL", "#4A148C"); // Z (Position Size only)
  styleGroup("AA1:AB1", "NOTES", "#212121"); // AA-AB (Tech + Fund Notes)
  styleGroup("AC1:AH1", "ENHANCED PATTERNS", "#6A1B9A"); // AC-AH (Enhanced indicators)

  calc.getRange("AE1")
    .setValue(syncTime)
    .setBackground("#000000")
    .setFontColor("#00FF00")
    .setFontWeight("bold")
    .setFontSize(9)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");

  // ------------------------------------------------------------------
  // ROW 2: COLUMN HEADERS (34 columns total - includes enhanced indicators AC-AH)
  // ------------------------------------------------------------------
  const headers = [[
    "Ticker", "SIGNAL", "FUNDAMENTAL", "DECISION", "Price", "Change %", "Vol Trend", "ATH (TRUE)", "ATH Diff %", "R:R Quality",
    "Trend Score", "Trend State", "SMA 20", "SMA 50", "SMA 200", "RSI", "MACD Hist", "Divergence", "ADX (14)", "Stoch %K (14)",
    "Support", "Resistance", "Target (3:1)", "ATR (14)", "Bollinger %B", "POSITION SIZE", "TECH NOTES", "FUND NOTES",
    "VOL REGIME", "ATH ZONE", "BBP SIGNAL", "PATTERNS", "ATR STOP", "ATR TARGET"
  ]];

  calc.getRange(2, 1, 1, 34)
    .setValues(headers)
    .setBackground("#111111")
    .setFontColor("white")
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setWrap(true);

  // Write tickers in A3:A
  if (tickers.length > 0) {
    calc.getRange(3, 1, tickers.length, 1).setValues(tickers.map(t => [t]));
  }

  // ------------------------------------------------------------------
  // FORMULAS
  // ------------------------------------------------------------------
  const formulas = [];

  const BLOCK = 7; // DATA block width (must match generateDataSheet)

  tickers.forEach((ticker, i) => {
    const row = i + 3;
    const t = String(ticker || "").trim().toUpperCase();

    // DATA block start (each ticker is 7 cols in DATA)
    const tDS = (i * BLOCK) + 1; // colStart
    const dateCol = columnToLetter(tDS + 0); // Date (row 5+)
    const openCol = columnToLetter(tDS + 1); // Open
    const highCol = columnToLetter(tDS + 2); // High
    const lowCol = columnToLetter(tDS + 3); // Low
    const closeCol = columnToLetter(tDS + 4); // Close
    const volCol = columnToLetter(tDS + 5); // Volume

    // Cached fundamentals in DATA row 3 (within same block)
    const athCell = `DATA!${columnToLetter(tDS + 1)}3`; // ATH value at colStart+1
    const peCell = `DATA!${columnToLetter(tDS + 3)}3`; // P/E value at colStart+3
    const epsCell = `DATA!${columnToLetter(tDS + 5)}3`; // EPS value at colStart+5

    // Rolling window anchors (row 5+ only)
    const lastRowCount = `COUNTA(DATA!$${closeCol}$5:$${closeCol})`; // number of data rows
    const lastAbsRow = `(4+${lastRowCount})`;                      // absolute row index
    const lastRowFormula = "COUNTA(DATA!$A:$A)";                      //used for support /resistence , to stay live

    // SIGNAL (B) — locale-safe + row5-anchored windows

    const useLongTermSignal =
      SpreadsheetApp.getActive().getSheetByName('INPUT').getRange('E2').getValue() === true;

    // ENHANCED SIGNAL LOGIC - Industry Standard + Pattern Recognition
    const fSignalLong =
      `=IF(OR(ISBLANK($E${row})${SEP}$E${row}=0)${SEP}"LOADING"${SEP}` +
      `IFS(` +
      // STOP OUT - Price below support level
      `$E${row}<$U${row}${SEP}"STOP OUT"${SEP}` +
      
      // RISK OFF - Price below long-term trend (SMA200)
      `$E${row}<$O${row}${SEP}"RISK OFF"${SEP}` +
      
      // ENHANCED PATTERN-BASED SIGNALS
      // ATH Breakout - New highs with volume and momentum
      `AND($I${row}>=-0.01${SEP}$G${row}>=1.5${SEP}$S${row}>=20${SEP}$E${row}>$O${row})${SEP}"ATH BREAKOUT"${SEP}` +
      
      // Volatility Breakout - ATR expansion with volume (fixed OFFSET bounds)
      `AND($X${row}>IFERROR(AVERAGE(OFFSET($X${row}${SEP}-MIN(20${SEP}ROW($X${row})-1)${SEP}0${SEP}MIN(20${SEP}ROW($X${row})-1)))${SEP}$X${row})*1.5${SEP}$G${row}>=2.0${SEP}$E${row}>$V${row})${SEP}"VOLATILITY BREAKOUT"${SEP}` +
      
      // Extreme Oversold - Multiple indicators aligned
      `AND($Y${row}<=0.1${SEP}$P${row}<=25${SEP}$T${row}<=0.20${SEP}$E${row}>$O${row})${SEP}"EXTREME OVERSOLD BUY"${SEP}` +
      
      // STRONG BUY - All bullish conditions aligned
      `AND(` +
        `$E${row}>$O${row}${SEP}` +        // Above SMA200
        `$N${row}>$O${row}${SEP}` +        // SMA50 > SMA200 (uptrend structure)
        `$P${row}<=30${SEP}` +             // RSI oversold
        `$Q${row}>0${SEP}` +               // MACD positive
        `$S${row}>=20${SEP}` +             // Strong trend (ADX)
        `$G${row}>=1.5` +                  // High volume confirmation
      `)${SEP}"STRONG BUY"${SEP}` +
      
      // BUY - Good entry conditions
      `AND(` +
        `$E${row}>$O${row}${SEP}` +        // Above SMA200
        `$N${row}>$O${row}${SEP}` +        // SMA50 > SMA200
        `$P${row}<=40${SEP}` +             // RSI not overbought
        `$Q${row}>0${SEP}` +               // MACD positive
        `$S${row}>=15` +                   // Trending
      `)${SEP}"BUY"${SEP}` +
      
      // ACCUMULATE - Dip buying opportunity
      `AND(` +
        `$E${row}>$O${row}${SEP}` +        // Above SMA200 (long-term uptrend)
        `$P${row}<=35${SEP}` +             // RSI pullback
        `$E${row}>=$N${row}*0.95` +        // Not too far from SMA50
      `)${SEP}"ACCUMULATE"${SEP}` +
      
      // OVERSOLD - Potential bounce
      `$P${row}<=20${SEP}"OVERSOLD"${SEP}` +
      
      // OVERBOUGHT - Potential pullback (enhanced with BBP)
      `OR($P${row}>=80${SEP}$Y${row}>=0.9)${SEP}"OVERBOUGHT"${SEP}` +
      
      // HOLD - Neutral conditions
      `AND(` +
        `$E${row}>$O${row}${SEP}` +        // Above SMA200
        `$P${row}>40${SEP}` +              // RSI above 40
        `$P${row}<70` +                    // RSI below 70
      `)${SEP}"HOLD"${SEP}` +
      
      // NEUTRAL - Default state
      `TRUE${SEP}"NEUTRAL"` +
      `)` +
      `)`;

    const fSignalTrend =
      `=IF(OR(ISBLANK($E${row})${SEP}$E${row}=0)${SEP}"LOADING"${SEP}` +
      `IFS(` +
      // STOP OUT - Price below support
      `$E${row}<$U${row}${SEP}"STOP OUT"${SEP}` +
      
      // RISK OFF - Price below long-term trend
      `$E${row}<$O${row}${SEP}"RISK OFF"${SEP}` +
      
      // ENHANCED PATTERN-BASED SIGNALS
      // Volatility Breakout - High priority (fixed OFFSET bounds)
      `AND($X${row}>IFERROR(AVERAGE(OFFSET($X${row}${SEP}-MIN(20${SEP}ROW($X${row})-1)${SEP}0${SEP}MIN(20${SEP}ROW($X${row})-1)))${SEP}$X${row})*1.5${SEP}$G${row}>=2.0${SEP}$E${row}>$V${row})${SEP}"VOLATILITY BREAKOUT"${SEP}` +
      
      // ATH Breakout - New highs with momentum
      `AND($I${row}>=-0.01${SEP}$G${row}>=1.5${SEP}$S${row}>=20)${SEP}"ATH BREAKOUT"${SEP}` +
      
      // BREAKOUT - High volume breakout above resistance
      `AND($G${row}>=1.5${SEP}$E${row}>=$V${row}*0.995)${SEP}"BREAKOUT"${SEP}` +
      
      // MOMENTUM - Strong trending conditions
      `AND($E${row}>$O${row}${SEP}$Q${row}>0${SEP}$S${row}>=20)${SEP}"MOMENTUM"${SEP}` +
      
      // UPTREND - Basic uptrend structure
      `AND($E${row}>$O${row}${SEP}$N${row}>$O${row}${SEP}$S${row}>=15)${SEP}"UPTREND"${SEP}` +
      
      // BULLISH - Above key moving averages
      `AND($E${row}>$N${row}${SEP}$E${row}>$M${row})${SEP}"BULLISH"${SEP}` +
      
      // OVERSOLD - Enhanced with BBP confirmation
      `AND(OR($T${row}<=0.20${SEP}$Y${row}<=0.2)${SEP}$E${row}>$U${row})${SEP}"OVERSOLD"${SEP}` +
      
      // OVERBOUGHT - Enhanced with BBP confirmation
      `OR($P${row}>=80${SEP}$Y${row}>=0.9)${SEP}"OVERBOUGHT"${SEP}` +
      
      // VOLATILITY SQUEEZE - Coiling pattern (fixed OFFSET bounds)
      `AND($X${row}<IFERROR(AVERAGE(OFFSET($X${row}${SEP}-MIN(20${SEP}ROW($X${row})-1)${SEP}0${SEP}MIN(20${SEP}ROW($X${row})-1)))${SEP}$X${row})*0.7${SEP}$S${row}<15${SEP}ABS($Y${row}-0.5)<0.2)${SEP}"VOLATILITY SQUEEZE"${SEP}` +
      
      // RANGE - Low trend strength
      `$S${row}<15${SEP}"RANGE"${SEP}` +
      
      // NEUTRAL - Default state
      `TRUE${SEP}"NEUTRAL"` +
      `)` +
      `)`;

    const fSignal = useLongTermSignal ? fSignalLong : fSignalTrend;

    // ENHANCED POSITION SIZING - ATR & ATH Risk Adjusted
    const fPositionSize = 
      `=IF($A${row}=""${SEP}""${SEP}` +
      `LET(` +
      `riskReward${SEP}$J${row}${SEP}` +
      `atrRisk${SEP}$X${row}/$E${row}${SEP}` +
      `athRisk${SEP}IF($I${row}>=-0.05${SEP}0.8${SEP}1.0)${SEP}` +        // Reduce size near ATH
      `volRegimeRisk${SEP}IFS(` +
        `atrRisk<=0.02${SEP}1.2${SEP}` +                                   // Low vol = larger size
        `atrRisk<=0.05${SEP}1.0${SEP}` +                                   // Normal vol = base size
        `atrRisk<=0.08${SEP}0.7${SEP}` +                                   // High vol = smaller size
        `TRUE${SEP}0.5` +                                                  // Extreme vol = much smaller
      `)${SEP}` +
      
      // Base size with risk adjustments
      `baseSize${SEP}0.02${SEP}` + // 2% base position
      `rrMultiplier${SEP}IF(riskReward>=3${SEP}1.5${SEP}IF(riskReward>=2${SEP}1.0${SEP}0.5))${SEP}` +
      
      `finalSize${SEP}MIN(0.08${SEP}baseSize*rrMultiplier*volRegimeRisk*athRisk)${SEP}` + // Max 8% position
      `TEXT(finalSize${SEP}"0.0%")&" (Vol: "&IFS(atrRisk<=0.02${SEP}"LOW"${SEP}atrRisk<=0.05${SEP}"NORM"${SEP}atrRisk<=0.08${SEP}"HIGH"${SEP}TRUE${SEP}"EXTR")&")"` +
      `)` +
      `)`;

    // DECISION (C) — unchanged gating pattern (kept stable)
    const tagExpr =
      `UPPER(IFERROR(INDEX(INPUT!$C$3:$C${SEP}MATCH($A${row}${SEP}INPUT!$A$3:$A${SEP}0))${SEP}"" ))`;

    const purchasedExpr = `ISNUMBER(SEARCH("PURCHASED"${SEP}${tagExpr}))`;

    const fDecisionLong =
      `=IF($A${row}=""${SEP}""${SEP}` +
      `IF($B${row}="LOADING"${SEP}"LOADING"${SEP}` +

      // PURCHASED => position management mode
      `IF(${purchasedExpr}${SEP}` +
      `IFS(` +
      `OR($B${row}="STOP OUT"${SEP}$B${row}="RISK OFF")${SEP}"🔴 EXIT"${SEP}` +

      `AND(OR($B${row}="STRONG BUY"${SEP}$B${row}="BUY"${SEP}$B${row}="ACCUMULATE")${SEP}OR(ISNUMBER(SEARCH("VALUE"${SEP}UPPER($C${row})))${SEP}ISNUMBER(SEARCH("FAIR"${SEP}UPPER($C${row})))))${SEP}` +
      `"🟢 ADD"${SEP}` +
      `AND(OR($B${row}="STRONG BUY"${SEP}$B${row}="BUY"${SEP}$B${row}="ACCUMULATE")${SEP}ISNUMBER(SEARCH("EXPENSIVE"${SEP}UPPER($C${row}))))${SEP}` +
      `"� HOLD / ADD SMALL"${SEP}` +
      `AND(OR($B${row}="STRONG BUY"${SEP}$B${row}="BUY"${SEP}$B${row}="ACCUMULATE")${SEP}ISNUMBER(SEARCH("PERFECTION"${SEP}UPPER($C${row}))))${SEP}` +
      `"🟡 HOLD (NO ADD)"${SEP}` +

      `AND($B${row}="OVERBOUGHT"${SEP}OR(ISNUMBER(SEARCH("EXPENSIVE"${SEP}UPPER($C${row})))${SEP}ISNUMBER(SEARCH("PERFECTION"${SEP}UPPER($C${row})))))${SEP}` +
      `"🟠 TRIM"${SEP}` +
      `$B${row}="HOLD"${SEP}"⚖️ HOLD"${SEP}` +

      `TRUE${SEP}"⚖️ HOLD"` +
      `)` +

      // NOT PURCHASED => entry mode
      `${SEP}` +
      `IFS(` +
      `OR($B${row}="STOP OUT"${SEP}$B${row}="RISK OFF")${SEP}"🔴 AVOID"${SEP}` +
      `$B${row}="STRONG BUY"${SEP}"🟢 STRONG BUY"${SEP}` +
      `OR($B${row}="BUY"${SEP}$B${row}="ACCUMULATE")${SEP}"🟢 BUY"${SEP}` +
      `$B${row}="OVERSOLD"${SEP}"🟡 WATCH (OVERSOLD)"${SEP}` +
      `$B${row}="OVERBOUGHT"${SEP}"⏳ WAIT (OVERBOUGHT)"${SEP}` +
      `$B${row}="HOLD"${SEP}"⚖️ WATCH"${SEP}` +
      `TRUE${SEP}"⚪ NEUTRAL"` +
      `)` +
      `)` +
      `)` +
      `)`;

    const fDecisionTrade =
      `=IF($A${row}=""${SEP}""${SEP}` +
      `LET(` +
      `tag${SEP}UPPER(IFERROR(INDEX(INPUT!$C$3:$C${SEP}MATCH($A${row}${SEP}INPUT!$A$3:$A${SEP}0))${SEP}"" ))${SEP}` +
      `purchased${SEP}REGEXMATCH(tag${SEP}"(^|,|\\s)PURCHASED(\\s|,|$)")${SEP}` +

      `IFS(` +
      // Stop out conditions
      `AND(IFERROR(VALUE($E${row})${SEP}0)>0${SEP}IFERROR(VALUE($U${row})${SEP}0)>0${SEP}IFERROR(VALUE($E${row})${SEP}0)<IFERROR(VALUE($U${row})${SEP}0))${SEP}` +
      `"Stop-Out"${SEP}` +

      // Take profit conditions (enhanced)
      `AND(purchased${SEP}OR($B${row}="OVERBOUGHT"${SEP}IFERROR(VALUE($E${row})${SEP}0)>=IFERROR(VALUE($W${row})${SEP}0)))${SEP}"Take Profit"${SEP}` +

      // Risk off conditions
      `AND(purchased${SEP}$B${row}="RISK OFF")${SEP}"Risk-Off"${SEP}` +
      `AND(NOT(purchased)${SEP}$B${row}="RISK OFF")${SEP}"Avoid"${SEP}` +

      // Enhanced entry conditions for non-purchased
      `AND(NOT(purchased)${SEP}OR($B${row}="VOLATILITY BREAKOUT"${SEP}$B${row}="ATH BREAKOUT")${SEP}OR($C${row}="VALUE"${SEP}$C${row}="FAIR"))${SEP}"Strong Trade Long"${SEP}` +
      `AND(NOT(purchased)${SEP}$B${row}="BREAKOUT"${SEP}OR($C${row}="VALUE"${SEP}$C${row}="FAIR"))${SEP}"Trade Long"${SEP}` +
      `AND(NOT(purchased)${SEP}$B${row}="MOMENTUM"${SEP}$C${row}="VALUE")${SEP}"Accumulate"${SEP}` +
      `AND(NOT(purchased)${SEP}$B${row}="OVERSOLD")${SEP}"Add in Dip"${SEP}` +
      `AND(NOT(purchased)${SEP}$B${row}="VOLATILITY SQUEEZE")${SEP}"Wait for Breakout"${SEP}` +

      // Hold conditions
      `$B${row}="MOMENTUM"${SEP}"Hold"${SEP}` +
      `$B${row}="UPTREND"${SEP}"Hold"${SEP}` +
      `$B${row}="BULLISH"${SEP}"Hold"${SEP}` +
      `TRUE${SEP}"Hold"` +
      `)` +
      `)` +
      `)`;

    const fDecision = useLongTermSignal ? fDecisionLong : fDecisionTrade;

    // Market regime is now referenced directly from DATA sheet in other formulas
    // No separate column needed - removed fMarketRegime

    // FUNDAMENTAL (D) — reads cached PE/EPS from DATA row 3 (fast)
    const fFund =
      `=IFERROR(` +
      `LET(` +
      `peRaw${SEP}${peCell}${SEP}` +
      `epsRaw${SEP}${epsCell}${SEP}` +
      `athDiffRaw${SEP}$I${row}${SEP}` +  // ATH Diff % column I

      `pe${SEP}IFERROR(VALUE(REGEXREPLACE(TO_TEXT(peRaw)${SEP}"[^0-9\\.\\-]"${SEP}"" ))${SEP}"" )${SEP}` +
      `eps${SEP}IFERROR(VALUE(REGEXREPLACE(TO_TEXT(epsRaw)${SEP}"[^0-9\\.\\-]"${SEP}"" ))${SEP}"" )${SEP}` +

      // Column I is % (e.g., -14.95%). Normalize to decimal (-0.1495).
      `athDiff${SEP}IFERROR(VALUE(REGEXREPLACE(TO_TEXT(athDiffRaw)${SEP}"[^0-9\\.\\-]"${SEP}"" ))/100${SEP}"" )${SEP}` +

      `IFS(` +
      `OR(pe=""${SEP}eps="")${SEP}"FAIR"${SEP}` +
      `eps<=0${SEP}"ZOMBIE"${SEP}` +

      // Priced for perfection = very high PE AND near ATH (within ~8%)
      `AND(pe>=60${SEP}athDiff<>""${SEP}athDiff>=-0.08)${SEP}"PRICED FOR PERFECTION"${SEP}` +

      `pe>=35${SEP}"EXPENSIVE"${SEP}` +
      `AND(pe>0${SEP}pe<=25${SEP}eps>=0.5)${SEP}"VALUE"${SEP}` +
      `AND(pe>25${SEP}pe<35${SEP}eps>=0.5)${SEP}"FAIR"${SEP}` +
      `TRUE${SEP}"FAIR"` +
      `)` + `)` + `${SEP}"FAIR")`;


    // E..Y
    const fPrice = `=ROUND(IFERROR(GOOGLEFINANCE("${t}"${SEP}"price")${SEP}0)${SEP}2)`;
    const fChg = `=IFERROR(GOOGLEFINANCE("${t}"${SEP}"changepct")/100${SEP}0)`;

    const fRVOL =
      `=ROUND(` +
      `IFERROR(` +
      `OFFSET(DATA!$${volCol}$5${SEP}${lastRowCount}-1${SEP}0)` +
      `/AVERAGE(OFFSET(DATA!$${volCol}$5${SEP}${lastRowCount}-20${SEP}0${SEP}20))` +
      `${SEP}1)` +
      `${SEP}2)`;

    const fATH = `=IFERROR(${athCell}${SEP}0)`;
    const fATHPct = `=IFERROR(($E${row}-$H${row})/MAX(0.01${SEP}$H${row})${SEP}0)`;

    const fRR =
      `=IF(OR($E${row}<=$U${row}${SEP}$E${row}=0)${SEP}0${SEP}` +
      `ROUND(MAX(0${SEP}$V${row}-$E${row})/MAX($X${row}*0.5${SEP}$E${row}-$U${row})${SEP}2)` +
      `)`;

    const fStars = `=REPT("★"${SEP} ($E${row}>$M${row}) + ($E${row}>$N${row}) + ($E${row}>$O${row}))`;
    const fTrend = `=IF($E${row}>$O${row}${SEP}"BULL"${SEP}"BEAR")`;

    const fSMA20 = `=ROUND(IFERROR(AVERAGE(OFFSET(DATA!$${closeCol}$5${SEP}${lastRowCount}-20${SEP}0${SEP}20))${SEP}0)${SEP}2)`;
    const fSMA50 = `=ROUND(IFERROR(AVERAGE(OFFSET(DATA!$${closeCol}$5${SEP}${lastRowCount}-50${SEP}0${SEP}50))${SEP}0)${SEP}2)`;
    const fSMA200 = `=ROUND(IFERROR(AVERAGE(OFFSET(DATA!$${closeCol}$5${SEP}${lastRowCount}-200${SEP}0${SEP}200))${SEP}0)${SEP}2)`;

    const fRSI = `=LIVERSI(DATA!$${closeCol}$5:$${closeCol}${SEP}$E${row})`;
    const fMACD = `=LIVEMACD(DATA!$${closeCol}$5:$${closeCol}${SEP}$E${row})`;

    const fDiv =
      `=IFERROR(IFS(` +
      `AND($E${row}<INDEX(DATA!$${closeCol}:$${closeCol}${SEP}${lastAbsRow}-14)${SEP}$P${row}>50)${SEP}"BULL DIV"${SEP}` +
      `AND($E${row}>INDEX(DATA!$${closeCol}:$${closeCol}${SEP}${lastAbsRow}-14)${SEP}$P${row}<50)${SEP}"BEAR DIV"${SEP}` +
      `TRUE${SEP}"—")${SEP}"—")`;

    const fADX = `=IFERROR(LIVEADX(DATA!$${highCol}$5:$${highCol}${SEP}DATA!$${lowCol}$5:$${lowCol}${SEP}DATA!$${closeCol}$5:$${closeCol}${SEP}$E${row})${SEP}0)`;
    const fStoch = `=LIVESTOCHK(DATA!$${highCol}$5:$${highCol}${SEP}DATA!$${lowCol}$5:$${lowCol}${SEP}DATA!$${closeCol}$5:$${closeCol}${SEP}$E${row})`;

    /**
     * Why this is the correct "Industry" fix:FeatureAverage of Extremes (Trial & Error)Percentile (Industry Standard)Outlier HandlingStill weighted by the outlier (e.g., 362.70).Ignores the outlier entirely.Zone AccuracyRepresents a single point.Represents the Value Area where most trading occurred.StabilityJumps around when a new high/low enters the window.Remains stable as long as the distribution of price is consistent.
     */
    const fRes = `=ROUND(IFERROR(LET(win${SEP}IFS($S${row}<20${SEP}10${SEP}$S${row}<35${SEP}22${SEP}TRUE${SEP}40)${SEP}n${SEP}${lastRowCount}${SEP}start${SEP}MAX(0${SEP}n-win)${SEP}len${SEP}MIN(win${SEP}n)${SEP}rng${SEP}IF(len<=0${SEP}OFFSET(DATA!$${highCol}$5${SEP}0${SEP}0)${SEP}OFFSET(DATA!$${highCol}$5${SEP}start${SEP}0${SEP}len))${SEP}out${SEP}IF(COUNTA(rng)<3${SEP}IFERROR(MAX(rng)${SEP}0)${SEP}PERCENTILE.INC(rng${SEP}0.85))${SEP}out)${SEP}0)${SEP}2)`;

    const fSup = `=ROUND(IFERROR(LET(win${SEP}IFS($S${row}<20${SEP}10${SEP}$S${row}<35${SEP}22${SEP}TRUE${SEP}40)${SEP}n${SEP}${lastRowCount}${SEP}start${SEP}MAX(0${SEP}n-win)${SEP}len${SEP}MIN(win${SEP}n)${SEP}rng${SEP}IF(len<=0${SEP}OFFSET(DATA!$${lowCol}$5${SEP}0${SEP}0)${SEP}OFFSET(DATA!$${lowCol}$5${SEP}start${SEP}0${SEP}len))${SEP}out${SEP}IF(COUNTA(rng)<3${SEP}IFERROR(MIN(rng)${SEP}0)${SEP}PERCENTILE.INC(rng${SEP}0.15))${SEP}out)${SEP}0)${SEP}2)`;

    // Target: Hybrid Logic (High of Resistance vs. 3:1 Projection)
    const fTgt = `=ROUND(MAX($V${row}${SEP}$E${row}+(($E${row}-$U${row})*3))${SEP}2)`;

    const fATR = `=IFERROR(LIVEATR(DATA!$${highCol}$5:$${highCol}${SEP}DATA!$${lowCol}$5:$${lowCol}${SEP}DATA!$${closeCol}$5:$${closeCol}${SEP}$E${row})${SEP}0)`;

    const fBBP =
      `=ROUND(IFERROR((($E${row}-$M${row})/(4*STDEV(OFFSET(DATA!$${closeCol}$5${SEP}${lastRowCount}-20${SEP}0${SEP}20))))+0.5${SEP}0.5)${SEP}2)`;

    // ENHANCED PATTERN RECOGNITION INDICATORS
    
    // ATR-based volatility regime
    const fVolRegime = 
      `=IFS(` +
      `$X${row}/$E${row}<=0.02${SEP}"LOW VOL"${SEP}` +
      `$X${row}/$E${row}<=0.05${SEP}"NORMAL VOL"${SEP}` +
      `$X${row}/$E${row}<=0.08${SEP}"HIGH VOL"${SEP}` +
      `TRUE${SEP}"EXTREME VOL"` +
      `)`;

    // ATH-based psychological zones
    const fATHZone = 
      `=IFS(` +
      `$I${row}>=-0.02${SEP}"AT ATH"${SEP}` +
      `$I${row}>=-0.05${SEP}"NEAR ATH"${SEP}` +
      `$I${row}>=-0.15${SEP}"RESISTANCE ZONE"${SEP}` +
      `$I${row}>=-0.30${SEP}"PULLBACK ZONE"${SEP}` +
      `$I${row}>=-0.50${SEP}"CORRECTION ZONE"${SEP}` +
      `TRUE${SEP}"DEEP VALUE ZONE"` +
      `)`;

    // BBP-based mean reversion signals
    const fBBSignal = 
      `=IFS(` +
      `AND($Y${row}>=0.9${SEP}$P${row}>=70)${SEP}"EXTREME OVERBOUGHT"${SEP}` +
      `AND($Y${row}<=0.1${SEP}$P${row}<=30)${SEP}"EXTREME OVERSOLD"${SEP}` +
      `AND($Y${row}>=0.8${SEP}$E${row}>$O${row})${SEP}"MOMENTUM STRONG"${SEP}` +
      `AND($Y${row}<=0.2${SEP}$E${row}>$U${row})${SEP}"MEAN REVERSION"${SEP}` +
      `TRUE${SEP}"NEUTRAL"` +
      `)`;

    // Pattern detection (fixed OFFSET bounds)
    const fPatterns = 
      `=TEXTJOIN(" | "${SEP}TRUE${SEP}` +
      // Volatility breakout
      `IF(AND($X${row}>IFERROR(AVERAGE(OFFSET($X${row}${SEP}-MIN(20${SEP}ROW($X${row})-1)${SEP}0${SEP}MIN(20${SEP}ROW($X${row})-1)))${SEP}$X${row})*1.5${SEP}$G${row}>=2.0${SEP}$E${row}>$V${row})${SEP}"VOL BREAKOUT"${SEP}"")${SEP}` +
      // ATH breakout
      `IF(AND($I${row}>=-0.01${SEP}$G${row}>=1.5${SEP}$S${row}>=20)${SEP}"ATH BREAKOUT"${SEP}"")${SEP}` +
      // Mean reversion setup
      `IF(AND($Y${row}<=0.15${SEP}$P${row}<=25${SEP}$T${row}<=0.20${SEP}$E${row}>$O${row})${SEP}"MEAN REVERSION SETUP"${SEP}"")${SEP}` +
      // Volatility squeeze
      `IF(AND($X${row}<IFERROR(AVERAGE(OFFSET($X${row}${SEP}-MIN(20${SEP}ROW($X${row})-1)${SEP}0${SEP}MIN(20${SEP}ROW($X${row})-1)))${SEP}$X${row})*0.7${SEP}$S${row}<15${SEP}ABS($Y${row}-0.5)<0.2)${SEP}"VOLATILITY SQUEEZE"${SEP}"")` +
      `)`;

    // ATR-based dynamic stops and targets
    const fATRStop = `=ROUND(MAX($U${row}${SEP}$E${row}-($X${row}*2))${SEP}2)`;
    const fATRTarget = `=ROUND($E${row}+($X${row}*3)${SEP}2)`;

    // Z TECH NOTES — parse-safe + correct columns + Stoch shown as %
    const fTechNotes =
      `=IF($A${row}=""${SEP}""${SEP}` +
      `"VOL: RVOL "&TEXT(IFERROR(VALUE($G${row})${SEP}0)${SEP}"0.00")&"x; "&` +
      `IF(IFERROR(VALUE($G${row})${SEP}0)<1${SEP}"sub-average (weak sponsorship)."${SEP}"healthy participation.")&CHAR(10)&` +

      `"REGIME: Price "&TEXT(IFERROR(VALUE($E${row})${SEP}0)${SEP}"0.00")&" vs SMA200 "&` +
      `TEXT(IFERROR(VALUE($O${row})${SEP}0)${SEP}"0.00")&"; "&` +
      `IF(IFERROR(VALUE($E${row})${SEP}0)<IFERROR(VALUE($O${row})${SEP}0)${SEP}"risk-off below SMA200."${SEP}"risk-on above SMA200.")&CHAR(10)&` +

      `"VOL/STRETCH: ATR(14) "&TEXT(IFERROR(VALUE($X${row})${SEP}0)${SEP}"0.00")&"; stretch "&` +
      `IF(` +
      `OR(IFERROR(VALUE($X${row})${SEP}0)=0${SEP}IFERROR(VALUE($M${row})${SEP}0)=0)${SEP}` +
      `"—"${SEP}` +
      `TEXT((IFERROR(VALUE($E${row})${SEP}0)-IFERROR(VALUE($M${row})${SEP}0))/IFERROR(VALUE($X${row})${SEP}1)${SEP}"0.0")&"x ATR"` +
      `)&" (<= +/-2x)."&CHAR(10)&` +

      `"MOMENTUM: RSI(14) "&TEXT(IFERROR(VALUE($P${row})${SEP}0)${SEP}"0.0")&"; "&` +
      `IF(IFERROR(VALUE($P${row})${SEP}0)<40${SEP}"negative bias."${SEP}"constructive.")&` +
      `" MACD hist "&TEXT(IFERROR(VALUE($Q${row})${SEP}0)${SEP}"0.000")&"; "&` +
      `IF(IFERROR(VALUE($Q${row})${SEP}0)>0${SEP}"improving."${SEP}"weak.")&CHAR(10)&` +

      `"TREND: ADX(14) "&TEXT(IFERROR(VALUE($S${row})${SEP}0)${SEP}"0.0")&"; "&` +
      `IF(IFERROR(VALUE($S${row})${SEP}0)>=25${SEP}"strong."${SEP}"weak.")&` +
      `" Stoch %K "&TEXT(IFERROR(VALUE($T${row})${SEP}0)${SEP}"0.0%")&" — "&` +
      `IF(IFERROR(VALUE($T${row})${SEP}0)<=0.2${SEP}"oversold zone (mean-reversion potential)."${SEP}` +
      `IF(IFERROR(VALUE($T${row})${SEP}0)>=0.8${SEP}"overbought zone (pullback risk)."${SEP}"neutral range (no timing edge)."))&CHAR(10)&` +

      `"R:R: "&TEXT(IFERROR(VALUE($J${row})${SEP}0)${SEP}"0.00")&"x; "&` +
      `IF(IFERROR(VALUE($J${row})${SEP}0)>=3${SEP}"favorable."${SEP}"limited")` +
      `)`;

    // AB FUND NOTES — Plain English explanation of signal path and decision flow using live data
    const fFundNotesLong =
      `=IF($A${row}=""${SEP}""${SEP}` +
      `IF(OR(ISBLANK($E${row})${SEP}$E${row}=0)${SEP}` +
      `"LOADING DATA..."${SEP}` +

      `TEXTJOIN(CHAR(10)&CHAR(10)${SEP}TRUE${SEP}` +

      // 1) WHY THIS SIGNAL FIRED (detailed explanation with indicator values)
      `"🔍 SIGNAL ANALYSIS: "&$B${row}&CHAR(10)&` +
      `IFS(` +
      // Enhanced Pattern Signals
      `$B${row}="ATH BREAKOUT"${SEP}` +
      `"BREAKOUT CONFIRMED: Price "&TEXT($E${row}${SEP}"$0.00")&" is within 1% of ATH "&TEXT($H${row}${SEP}"$0.00")&" (difference: "&TEXT($I${row}${SEP}"+0.0%;-0.0%")&"). Volume at "&TEXT($G${row}${SEP}"0.0")&"x average confirms institutional participation. ADX "&TEXT($S${row}${SEP}"0.0")&" shows strong trend momentum (≥20 required). This combination signals professional accumulation at new highs."${SEP}` +
      
      `$B${row}="VOLATILITY BREAKOUT"${SEP}` +
      `"EXPLOSIVE EXPANSION: ATR "&TEXT($X${row}${SEP}"$0.00")&" expanded 50%+ above 20-period average, indicating volatility breakout. Volume "&TEXT($G${row}${SEP}"0.0")&"x normal (≥2.0x required) with price "&TEXT($E${row}${SEP}"$0.00")&" breaking resistance "&TEXT($V${row}${SEP}"$0.00")&". This pattern indicates institutional participation driving momentum expansion."${SEP}` +
      
      `$B${row}="EXTREME OVERSOLD BUY"${SEP}` +
      `"MULTI-INDICATOR OVERSOLD: Three oversold signals aligned - Bollinger %B "&TEXT($Y${row}${SEP}"0.0%")&" (≤10% = extreme), RSI "&TEXT($P${row}${SEP}"0.0")&" (≤25 = oversold), Stochastic "&TEXT($T${row}${SEP}"0.0%")&" (≤20% = oversold). Price "&TEXT($E${row}${SEP}"$0.00")&" remains above SMA200 "&TEXT($O${row}${SEP}"$0.00")&" confirming bullish regime. This creates high-probability bounce setup."${SEP}` +
      
      // Standard Signals with detailed explanations
      `$B${row}="STRONG BUY"${SEP}` +
      `"PERFECT STORM SETUP: All 6 bullish conditions met: (1) Price "&TEXT($E${row}${SEP}"$0.00")&" > SMA200 "&TEXT($O${row}${SEP}"$0.00")&" (bullish regime), (2) SMA50 "&TEXT($N${row}${SEP}"$0.00")&" > SMA200 (uptrend structure), (3) RSI "&TEXT($P${row}${SEP}"0.0")&" ≤30 (oversold entry), (4) MACD "&TEXT($Q${row}${SEP}"0.000")&" >0 (positive momentum), (5) ADX "&TEXT($S${row}${SEP}"0.0")&" ≥20 (strong trend), (6) Volume "&TEXT($G${row}${SEP}"0.0")&"x ≥1.5 (institutional participation). Rare alignment for institutional buying."${SEP}` +
      
      `$B${row}="BUY"${SEP}` +
      `"SOLID ENTRY SETUP: Key conditions met: Price "&TEXT($E${row}${SEP}"$0.00")&" > SMA200 "&TEXT($O${row}${SEP}"$0.00")&" (bullish regime), SMA50 "&TEXT($N${row}${SEP}"$0.00")&" > SMA200 (uptrend structure), RSI "&TEXT($P${row}${SEP}"0.0")&" ≤40 (not overbought), MACD "&TEXT($Q${row}${SEP}"0.000")&" >0 (positive momentum), ADX "&TEXT($S${row}${SEP}"0.0")&" ≥15 (trending). Good institutional accumulation zone without oversold requirement."${SEP}` +
      
      `$B${row}="ACCUMULATE"${SEP}` +
      `"DIP BUYING OPPORTUNITY: Bullish regime confirmed with Price "&TEXT($E${row}${SEP}"$0.00")&" > SMA200 "&TEXT($O${row}${SEP}"$0.00")&". Pullback conditions: RSI "&TEXT($P${row}${SEP}"0.0")&" ≤35 (pullback level) and price "&TEXT($E${row}${SEP}"$0.00")&" near SMA50 "&TEXT($N${row}${SEP}"$0.00")&" (≥95% of SMA50 required). Classic institutional accumulation pattern - buying weakness in established uptrend."${SEP}` +
      
      // Risk Signals with clear explanations
      `$B${row}="STOP OUT"${SEP}` +
      `"CRITICAL BREAKDOWN: Price "&TEXT($E${row}${SEP}"$0.00")&" broke below support "&TEXT($U${row}${SEP}"$0.00")&". This violates the bullish structure and triggers immediate risk management exit. Support breaks indicate institutional distribution and potential trend reversal - capital preservation overrides all other factors."${SEP}` +
      
      `$B${row}="RISK OFF"${SEP}` +
      `"BEARISH REGIME CONFIRMED: Price "&TEXT($E${row}${SEP}"$0.00")&" fell below SMA200 "&TEXT($O${row}${SEP}"$0.00")&". The 200-day moving average is the primary trend filter - below indicates long-term bearish regime. Institutional money typically exits when this level breaks. Avoid new positions until price reclaims SMA200."${SEP}` +
      
      `$B${row}="OVERBOUGHT"${SEP}` +
      `"PULLBACK WARNING: "&IF($P${row}>=80${SEP}"RSI "&TEXT($P${row}${SEP}"0.0")&" ≥80 (overbought threshold)"${SEP}"")&IF(AND($P${row}>=80${SEP}$Y${row}>=0.9)${SEP}" and "${SEP}"")&IF($Y${row}>=0.9${SEP}"Bollinger %B "&TEXT($Y${row}${SEP}"0.0%")&" ≥90% (upper extreme)"${SEP}"")&". Price "&TEXT($E${row}${SEP}"$0.00")&" vulnerable to profit-taking as momentum indicators reach extreme levels."${SEP}` +
      
      `$B${row}="OVERSOLD"${SEP}` +
      `"BOUNCE POTENTIAL: RSI "&TEXT($P${row}${SEP}"0.0")&" ≤20 indicates extreme oversold conditions. While this suggests potential bounce, need confirmation from price action and other indicators. Monitor for reversal signals before acting."${SEP}` +
      
      `$B${row}="HOLD"${SEP}` +
      `"NEUTRAL CONDITIONS: Price "&TEXT($E${row}${SEP}"$0.00")&" > SMA200 "&TEXT($O${row}${SEP}"$0.00")&" confirms bullish regime, but RSI "&TEXT($P${row}${SEP}"0.0")&" in neutral range (40-70). No clear momentum edge for new positions - maintain current exposure."${SEP}` +
      
      `$B${row}="NEUTRAL"${SEP}` +
      `"NO CLEAR SETUP: Mixed signals prevent clear direction. Price "&TEXT($E${row}${SEP}"$0.00")&" vs SMA200 "&TEXT($O${row}${SEP}"$0.00")&" shows "&IF($E${row}>$O${row}${SEP}"bullish regime but"${SEP}"bearish regime and")&" RSI "&TEXT($P${row}${SEP}"0.0")&", ADX "&TEXT($S${row}${SEP}"0.0")&" provide conflicting signals. Wait for clearer technical alignment."${SEP}` +
      
      `TRUE${SEP}` +
      `"STRUCTURE/MOMENTUM NOT BULLISH: Key indicators show weakness - "&IF($E${row}<$O${row}${SEP}"Price "&TEXT($E${row}${SEP}"$0.00")&" < SMA200 "&TEXT($O${row}${SEP}"$0.00")&" (bearish regime)"${SEP}"")&IF(AND($E${row}<$O${row}${SEP}$S${row}<15)${SEP}", "${SEP}"")&IF($S${row}<15${SEP}"ADX "&TEXT($S${row}${SEP}"0.0")&" <15 (no trend)"${SEP}"")&IF(AND($E${row}<$O${row}${SEP}$Q${row}<=0)${SEP}", "${SEP}IF(AND($E${row}>=$O${row}${SEP}$Q${row}<=0)${SEP}", "${SEP}""))&IF($Q${row}<=0${SEP}"MACD "&TEXT($Q${row}${SEP}"0.000")&" ≤0 (negative momentum)"${SEP}"")&". Without bullish structure and momentum alignment, position is watchlist only."` +
      `)${SEP}` +

      // 2) WHY THIS FUNDAMENTAL RATING (with specific P/E and EPS explanations)
      `"💰 VALUATION ANALYSIS: "&$C${row}&CHAR(10)&` +
      `IFS(` +
      `$C${row}="VALUE"${SEP}` +
      `"ATTRACTIVE VALUATION: P/E ratio "&TEXT(IFERROR(GOOGLEFINANCE($A${row}${SEP}"pe")${SEP}0)${SEP}"0.0")&" is reasonable (≤25 threshold) with positive EPS "&TEXT(IFERROR(GOOGLEFINANCE($A${row}${SEP}"eps")${SEP}0)${SEP}"$0.00")&" (≥$0.50 required). Company generates solid earnings relative to stock price, providing margin of safety. Fundamentals support long-term investment thesis."${SEP}` +
      
      `$C${row}="FAIR"${SEP}` +
      `"NEUTRAL VALUATION: P/E ratio "&TEXT(IFERROR(GOOGLEFINANCE($A${row}${SEP}"pe")${SEP}0)${SEP}"0.0")&" in fair range (25-35) with EPS "&TEXT(IFERROR(GOOGLEFINANCE($A${row}${SEP}"eps")${SEP}0)${SEP}"$0.00")&". Stock neither cheap nor expensive - trading near fair value. Valuation neutral to investment decision, focus on technical factors."${SEP}` +
      
      `$C${row}="EXPENSIVE"${SEP}` +
      `"PREMIUM VALUATION: P/E ratio "&TEXT(IFERROR(GOOGLEFINANCE($A${row}${SEP}"pe")${SEP}0)${SEP}"0.0")&" is elevated (35-60 range) relative to earnings "&TEXT(IFERROR(GOOGLEFINANCE($A${row}${SEP}"eps")${SEP}0)${SEP}"$0.00")&". Market has high expectations built into price. Creates valuation headwind - requires strong execution to justify current levels, less margin for error."${SEP}` +
      
      `$C${row}="PRICED FOR PERFECTION"${SEP}` +
      `"EXTREME VALUATION: P/E ratio "&TEXT(IFERROR(GOOGLEFINANCE($A${row}${SEP}"pe")${SEP}0)${SEP}"0.0")&" is very high (≥60) and price "&TEXT($E${row}${SEP}"$0.00")&" near ATH "&TEXT($H${row}${SEP}"$0.00")&" ("&TEXT($I${row}${SEP}"+0.0%;-0.0%")&" from highs). Market expects perfect execution - any disappointment in earnings, guidance, or growth could trigger significant selloff."${SEP}` +
      
      `$C${row}="ZOMBIE"${SEP}` +
      `"POOR FUNDAMENTALS: Company has negative or very weak EPS "&TEXT(IFERROR(GOOGLEFINANCE($A${row}${SEP}"eps")${SEP}0)${SEP}"$0.00")&" (≤$0 threshold). Business losing money or barely profitable. High risk of permanent capital loss - fundamentals severely negative, avoid unless purely technical momentum trade."${SEP}` +
      
      `TRUE${SEP}` +
      `"UNCLEAR FUNDAMENTALS: Missing P/E or EPS data prevents valuation assessment. Could indicate recent IPO, financial restatement, or data provider issues. Treat as neutral and focus on technical risk management."` +
      `)${SEP}` +

      // 3) WHY THIS DECISION (combining signal + fundamental + position status)
      `"⚡ INVESTMENT DECISION: "&$D${row}&CHAR(10)&` +
      `IFS(` +
      // Risk Management Decisions
      `ISNUMBER(SEARCH("EXIT"${SEP}$C${row}))${SEP}` +
      `"CAPITAL PRESERVATION OVERRIDE: Signal "&$B${row}&" triggered risk management rules that override all other factors including "&$C${row}&" fundamentals. When price breaks support ("&TEXT($U${row}${SEP}"$0.00")&") or trend breaks (SMA200 "&TEXT($O${row}${SEP}"$0.00")&"), institutional discipline requires immediate exit to preserve capital."${SEP}` +
      
      `ISNUMBER(SEARCH("AVOID"${SEP}$D${row}))${SEP}` +
      `"RISK-OFF REGIME: Signal "&$B${row}&" indicates bearish market conditions. Price below SMA200 suggests institutional distribution phase. No new positions until technical structure improves - focus capital on better opportunities."${SEP}` +
      
      // Buy Decisions
      `ISNUMBER(SEARCH("STRONG BUY"${SEP}$D${row}))${SEP}` +
      `"HIGH CONVICTION ENTRY: Signal "&$B${row}&" shows exceptional technical setup. "&IF(ISNUMBER(SEARCH("PURCHASED"${SEP}UPPER(IFERROR(INDEX(INPUT!$C$3:$C${SEP}MATCH($A${row}${SEP}INPUT!$A$3:$A${SEP}0))${SEP}""))))${SEP}"Adding to existing position"${SEP}"Initiating new position")&" with high confidence. Technical excellence combined with "&$C${row}&" valuation creates compelling risk/reward opportunity."${SEP}` +
      
      `ISNUMBER(SEARCH("BUY"${SEP}$D${row}))${SEP}` +
      `"SOLID ENTRY OPPORTUNITY: Signal "&$B${row}&" provides good technical setup with "&$C${row}&" fundamental backdrop. "&IF(ISNUMBER(SEARCH("PURCHASED"${SEP}UPPER(IFERROR(INDEX(INPUT!$C$3:$C${SEP}MATCH($A${row}${SEP}INPUT!$A$3:$A${SEP}0))${SEP}""))))${SEP}"Adding to position"${SEP}"Starting position")&" as technical structure and valuation align favorably for institutional accumulation."${SEP}` +
      
      `ISNUMBER(SEARCH("ADD"${SEP}$D${row}))${SEP}` +
      `"POSITION BUILDING: Signal "&$B${row}&" with "&$C${row}&" fundamentals supports scaling into position. Technical setup warrants adding to existing holdings while maintaining disciplined risk controls and position sizing."${SEP}` +
      
      // Hold/Wait Decisions
      `ISNUMBER(SEARCH("TRIM"${SEP}$D${row}))${SEP}` +
      `"PROFIT TAKING DISCIPLINE: Signal "&$B${row}&" shows overbought technical conditions while "&$C${row}&" valuation creates fundamental headwinds. Reducing position size to lock in gains and reduce exposure to pullback risk - disciplined profit taking."${SEP}` +
      
      `ISNUMBER(SEARCH("HOLD"${SEP}$D${row}))${SEP}` +
      `"MAINTAIN CURRENT STANCE: Signal "&$B${row}&" shows mixed technical conditions. "&IF(ISNUMBER(SEARCH("PURCHASED"${SEP}UPPER(IFERROR(INDEX(INPUT!$C$3:$C${SEP}MATCH($A${row}${SEP}INPUT!$A$3:$A${SEP}0))${SEP}""))))${SEP}"Holding existing position"${SEP}"Staying on sidelines")&" until clearer directional signals emerge from market structure."${SEP}` +
      
      `ISNUMBER(SEARCH("WAIT"${SEP}$C${row}))${SEP}` +
      `"PATIENCE FOR BETTER ENTRY: Signal "&$B${row}&" suggests waiting for improved risk/reward. Current setup at price "&TEXT($E${row}${SEP}"$0.00")&" vs target "&TEXT($W${row}${SEP}"$0.00")&" offers "&TEXT(($W${row}-$E${row})/$E${row}${SEP}"+0.0%;-0.0%")&" upside - not compelling enough given "&$D${row}&" valuation."${SEP}` +
      
      `ISNUMBER(SEARCH("WATCH"${SEP}$C${row}))${SEP}` +
      `"MONITORING FOR CONFIRMATION: Signal "&$B${row}&" shows potential but requires additional confirmation. Watching for follow-through in price action, volume, and momentum before committing capital."${SEP}` +
      
      `TRUE${SEP}` +
      `"NEUTRAL STANCE: Signal "&$B${row}&" combined with "&$D${row}&" fundamentals provides no clear investment edge. Maintaining current position until market conditions and technical setup provide clearer directional bias."` +
      `)` +

      `)` +  // TEXTJOIN
      `)` +
      `)`;

    const fFundNotesTrade =
      `=IF($A${row}=""${SEP}""${SEP}` +
      `IF(OR(ISBLANK($E${row})${SEP}$E${row}=0)${SEP}` +
      `"LOADING DATA..."${SEP}` +

      `TEXTJOIN(CHAR(10)&CHAR(10)${SEP}TRUE${SEP}` +

      // 1) WHY THIS SIGNAL FIRED (with specific indicator explanations)
      `"🔍 SIGNAL ANALYSIS: "&$B${row}&CHAR(10)&` +
      `IFS(` +
      // Enhanced Pattern Signals
      `$B${row}="VOLATILITY BREAKOUT"${SEP}` +
      `"EXPLOSIVE MOMENTUM: ATR "&TEXT($X${row}${SEP}"$0.00")&" expanded 50%+ above 20-period average (volatility breakout threshold). Volume "&TEXT($G${row}${SEP}"0.0")&"x normal (≥2.0x required) with price "&TEXT($E${row}${SEP}"$0.00")&" breaking resistance "&TEXT($V${row}${SEP}"$0.00")&". This pattern signals institutional participation driving momentum expansion."${SEP}` +
      
      `$B${row}="ATH BREAKOUT"${SEP}` +
      `"NEW HIGH MOMENTUM: Price "&TEXT($E${row}${SEP}"$0.00")&" within 1% of ATH "&TEXT($H${row}${SEP}"$0.00")&" with volume "&TEXT($G${row}${SEP}"0.0")&"x (≥1.5x required) and ADX "&TEXT($S${row}${SEP}"0.0")&" (≥20 trend strength). Breaking psychological resistance at new highs with institutional participation."${SEP}` +
      
      // Standard Tactical Signals
      `$B${row}="BREAKOUT"${SEP}` +
      `"RESISTANCE BREAKTHROUGH: Price "&TEXT($E${row}${SEP}"$0.00")&" broke above resistance "&TEXT($V${row}${SEP}"$0.00")&" with volume "&TEXT($G${row}${SEP}"0.0")&"x confirmation (≥1.5x threshold). Technical breakout suggests continuation of upward momentum."${SEP}` +
      
      `$B${row}="MOMENTUM"${SEP}` +
      `"TREND STRENGTH CONFIRMED: Price "&TEXT($E${row}${SEP}"$0.00")&" > SMA200 "&TEXT($O${row}${SEP}"$0.00")&" (bullish regime), MACD "&TEXT($Q${row}${SEP}"0.000")&" >0 (positive momentum), ADX "&TEXT($S${row}${SEP}"0.0")&" ≥20 (strong trend). All momentum indicators aligned for trend continuation."${SEP}` +
      
      `$B${row}="UPTREND"${SEP}` +
      `"BASIC TREND STRUCTURE: Price "&TEXT($E${row}${SEP}"$0.00")&" > SMA200 "&TEXT($O${row}${SEP}"$0.00")&", SMA50 "&TEXT($N${row}${SEP}"$0.00")&" > SMA200 (uptrend alignment), ADX "&TEXT($S${row}${SEP}"0.0")&" ≥15 (trending threshold). Basic uptrend structure intact for trend following."${SEP}` +
      
      `$B${row}="BULLISH"${SEP}` +
      `"SHORT-TERM BULLISH BIAS: Price "&TEXT($E${row}${SEP}"$0.00")&" above both SMA50 "&TEXT($N${row}${SEP}"$0.00")&" and SMA20 "&TEXT($M${row}${SEP}"$0.00")&". Near-term moving average alignment shows positive short-term bias."${SEP}` +
      
      // Mean Reversion
      `$B${row}="OVERSOLD"${SEP}` +
      `"BOUNCE SETUP: Stochastic "&TEXT($T${row}${SEP}"0.0%")&" ≤20% (oversold threshold) while price "&TEXT($E${row}${SEP}"$0.00")&" holds above support "&TEXT($U${row}${SEP}"$0.00")&". Oversold condition in uptrend creates potential counter-trend bounce opportunity."${SEP}` +
      
      `$B${row}="VOLATILITY SQUEEZE"${SEP}` +
      `"COILING FOR BREAKOUT: ATR "&TEXT($X${row}${SEP}"$0.00")&" compressed below 20-period average (low volatility), ADX "&TEXT($S${row}${SEP}"0.0")&" <15 (no trend), Bollinger %B "&TEXT($Y${row}${SEP}"0.0%")&" near 50% (midpoint). Volatility compression often precedes directional moves."${SEP}` +
      
      // Risk Management
      `$B${row}="STOP OUT"${SEP}` +
      `"SUPPORT BREAKDOWN: Price "&TEXT($E${row}${SEP}"$0.00")&" broke below support "&TEXT($U${row}${SEP}"$0.00")&". Support break invalidates bullish thesis and triggers defensive exit to preserve capital."${SEP}` +
      
      `$B${row}="RISK OFF"${SEP}` +
      `"BEARISH REGIME: Price "&TEXT($E${row}${SEP}"$0.00")&" below SMA200 "&TEXT($O${row}${SEP}"$0.00")&". Long-term trend broken indicates institutional distribution - avoid new long positions."${SEP}` +
      
      `$B${row}="OVERBOUGHT"${SEP}` +
      `"PULLBACK RISK: RSI "&TEXT($P${row}${SEP}"0.0")&" ≥80 (overbought threshold). Price vulnerable to profit-taking from elevated momentum levels."${SEP}` +
      
      `$B${row}="RANGE"${SEP}` +
      `"SIDEWAYS MARKET: ADX "&TEXT($S${row}${SEP}"0.0")&" <15 indicates no clear trend. Price chopping between levels - range trading tactics only, avoid directional bets."${SEP}` +
      
      `TRUE${SEP}` +
      `"MIXED SIGNALS: No clear technical setup. Price "&TEXT($E${row}${SEP}"$0.00")&", RSI "&TEXT($P${row}${SEP}"0.0")&", ADX "&TEXT($S${row}${SEP}"0.0")&" show conflicting signals - wait for clarity."` +
      `)${SEP}` +

      // 2) FUNDAMENTAL CONTEXT (concise for trade mode)
      `"💰 VALUATION CONTEXT: "&$D${row}&CHAR(10)&` +
      `IFS(` +
      `$D${row}="VALUE"${SEP}` +
      `"SUPPORTIVE: P/E "&TEXT(IFERROR(GOOGLEFINANCE($A${row}${SEP}"pe")${SEP}0)${SEP}"0.0")&" reasonable (≤25), EPS "&TEXT(IFERROR(GOOGLEFINANCE($A${row}${SEP}"eps")${SEP}0)${SEP}"$0.00")&" positive (≥$0.50). Attractive valuation provides fundamental tailwind for trade."${SEP}` +
      
      `$D${row}="FAIR"${SEP}` +
      `"NEUTRAL: P/E "&TEXT(IFERROR(GOOGLEFINANCE($A${row}${SEP}"pe")${SEP}0)${SEP}"0.0")&" in fair range (25-35). Valuation neither helps nor hurts the trade setup."${SEP}` +
      
      `$D${row}="EXPENSIVE"${SEP}` +
      `"HEADWIND: P/E "&TEXT(IFERROR(GOOGLEFINANCE($A${row}${SEP}"pe")${SEP}0)${SEP}"0.0")&" elevated (35-60). Premium valuation creates less margin for error - tighter risk controls needed."${SEP}` +
      
      `$D${row}="PRICED FOR PERFECTION"${SEP}` +
      `"FRAGILE: P/E "&TEXT(IFERROR(GOOGLEFINANCE($A${row}${SEP}"pe")${SEP}0)${SEP}"0.0")&" extreme (≥60), near ATH. High reversal risk on any disappointment - very tight stops required."${SEP}` +
      
      `$D${row}="ZOMBIE"${SEP}` +
      `"HIGH RISK: Negative EPS "&TEXT(IFERROR(GOOGLEFINANCE($A${row}${SEP}"eps")${SEP}0)${SEP}"$0.00")&" (≤$0). Pure momentum play with fundamental weakness - essential tight risk controls."${SEP}` +
      
      `TRUE${SEP}` +
      `"UNCLEAR: Missing fundamental data. Focus purely on technical factors and risk management."` +
      `)${SEP}` +

      // 3) DECISION LOGIC
      `"⚡ TRADE DECISION: "&$C${row}&CHAR(10)&` +
      `IFS(` +
      `$C${row}="Stop-Out"${SEP}` +
      `"EXIT REQUIRED: Support break at "&TEXT($U${row}${SEP}"$0.00")&" invalidates bullish setup. Cut losses immediately to preserve trading capital."${SEP}` +
      
      `$C${row}="Take Profit"${SEP}` +
      `"LOCK IN GAINS: "&IF(IFERROR(VALUE($E${row})${SEP}0)>=IFERROR(VALUE($W${row})${SEP}0)${SEP}"Target "&TEXT($W${row}${SEP}"$0.00")&" reached ("&TEXT(($E${row}-$W${row})/$W${row}${SEP}"+0.0%;-0.0%")&" vs target)"${SEP}"Overbought conditions with RSI "&TEXT($P${row}${SEP}"0.0"))&". Taking profits while momentum favorable."${SEP}` +
      
      `$C${row}="Strong Trade Long"${SEP}` +
      `"HIGH-CONVICTION ENTRY: "&$B${row}&" signal with "&$D${row}&" backdrop creates exceptional setup. Target "&TEXT($W${row}${SEP}"$0.00")&" ("&TEXT(($W${row}-$E${row})/$E${row}${SEP}"+0.0%;-0.0%")&" upside) vs stop "&TEXT($U${row}${SEP}"$0.00")&" offers strong risk/reward."${SEP}` +
      
      `$C${row}="Trade Long"${SEP}` +
      `"SOLID ENTRY: "&$B${row}&" setup with "&$D${row}&" valuation. Risk/reward ratio "&TEXT($J${row}${SEP}"0.0")&":1 supports position initiation."${SEP}` +
      
      `$C${row}="Accumulate"${SEP}` +
      `"BUILD POSITION: "&$B${row}&" in trending market with VALUE fundamental support. Adding to position on strength."${SEP}` +
      
      `$C${row}="Add in Dip"${SEP}` +
      `"DIP BUYING: Oversold bounce setup while holding above support. Adding to winning position on temporary weakness."${SEP}` +
      
      `$C${row}="Wait for Breakout"${SEP}` +
      `"PATIENCE REQUIRED: Volatility squeeze needs directional resolution. Wait for clear breakout direction before entering."${SEP}` +
      
      `$C${row}="Risk-Off"${SEP}` +
      `"DEFENSIVE MODE: Below SMA200 indicates bearish regime. Exit long positions, avoid new entries until trend improves."${SEP}` +
      
      `$C${row}="Avoid"${SEP}` +
      `"POOR SETUP: Unfavorable technical structure with weak risk/reward. Focus capital on better opportunities."${SEP}` +
      
      `TRUE${SEP}` +
      `"HOLD CURRENT: Mixed signals suggest maintaining current stance until technical clarity emerges."` +
      `)` +

      `)` +  // TEXTJOIN
      `)` +
      `)`;

    const fFundNotes = useLongTermSignal ? fFundNotesLong : fFundNotesTrade;


    formulas.push([
      fSignal,      // B
      fFund,        // C (swapped from D)
      fDecision,    // D (swapped from C)
      fPrice,       // E
      fChg,         // F
      fRVOL,        // G
      fATH,         // H
      fATHPct,      // I
      fRR,          // J
      fStars,       // K
      fTrend,       // L
      fSMA20,       // M
      fSMA50,       // N
      fSMA200,      // O
      fRSI,         // P
      fMACD,        // Q
      fDiv,         // R
      fADX,         // S
      fStoch,       // T
      fSup,         // U
      fRes,         // V
      fTgt,         // W
      fATR,         // X
      fBBP,         // Y
      fPositionSize, // Z - Enhanced position sizing
      fTechNotes,   // AA - Technical notes  
      fFundNotes,   // AB - Fundamental notes
      fVolRegime,   // AC - Volatility regime
      fATHZone,     // AD - ATH psychological zones
      fBBSignal,    // AE - BBP mean reversion signals
      fPatterns,    // AF - Pattern detection
      fATRStop,     // AG - ATR-based stop loss
      fATRTarget    // AH - ATR-based target
    ]);
  });

  if (tickers.length > 0) {
    // B..AH (33 cols total: B-Y=24, Z=1 position size, AA-AB=2 notes, AC-AH=6 enhanced indicators)
    calc.getRange(3, 2, formulas.length, 33).setFormulas(formulas);
  }

  // ------------------------------------------------------------------
  // FORMATTING (kept consistent with your current style)
  // ------------------------------------------------------------------
  const lr = Math.max(calc.getLastRow(), 3);
  calc.setFrozenRows(2);

  if (lr > 2) {
    const dataRows = lr - 2;
    calc.setRowHeights(3, dataRows, 72);
    calc.getRange(3, 1, dataRows, 34)
      .setHorizontalAlignment("left")
      .setVerticalAlignment("middle")
      .setWrap(true);
  }

  for (let c = 1; c <= 25; c++) calc.setColumnWidth(c, 90);
  calc.setColumnWidth(26, 100); // Z POSITION SIZE
  calc.setColumnWidth(27, 420); // AA TECH NOTES
  calc.setColumnWidth(28, 420); // AB FUND NOTES
  calc.setColumnWidth(29, 120); // AC VOL REGIME
  calc.setColumnWidth(30, 140); // AD ATH ZONE
  calc.setColumnWidth(31, 160); // AE BB SIGNAL
  calc.setColumnWidth(32, 300); // AF PATTERNS
  calc.setColumnWidth(33, 90);  // AG ATR STOP
  calc.setColumnWidth(34, 90);  // AH ATR TARGET

  calc.getRange("F3:F").setNumberFormat("0.00%");
  calc.getRange("I3:I").setNumberFormat("0.00%");
  calc.getRange("T3:T").setNumberFormat("0.00%"); // Stoch 0..1
  calc.getRange("Y3:Y").setNumberFormat("0.00%");

  const lastRowAll = Math.max(calc.getLastRow(), 2);
  calc.getRange(1, 1, lastRowAll, 34)
    .setBorder(true, true, true, true, true, true, "#000000", SpreadsheetApp.BorderStyle.SOLID);
  calc.getRange("A1:AH2")
    .setBorder(true, true, true, true, true, true, "#FFFFFF", SpreadsheetApp.BorderStyle.SOLID);

  SpreadsheetApp.flush();
}

function generateDashboardSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const input = ss.getSheetByName("INPUT");
  if (!input) return;

  const dashboard = ss.getSheetByName("DASHBOARD") || ss.insertSheet("DASHBOARD");

  const tickers = getCleanTickers(input);
  const DATA_START_ROW = 4;
  const DATA_ROWS = Math.max(50, Math.min(500, tickers.length + 40)); // spill cushion
  const SENTINEL = "DASHBOARD_LAYOUT_V1_BLOOMBERG";

  const isInitialized = (dashboard.getRange("A1").getNote() || "").indexOf(SENTINEL) !== -1;

  // ==========================
  // ONE-TIME LAYOUT (NO ROW3+ formatting during refresh)
  // ==========================
  if (!isInitialized) {
    dashboard.clear().clearFormats();

    // --- Row 1 controls ---
    dashboard.getRange("A1")
      .setValue("UPDATE CAL")
      .setBackground("#212121").setFontColor("white").setFontWeight("bold")
      .setHorizontalAlignment("center").setVerticalAlignment("middle");

    dashboard.getRange("B1")
      .insertCheckboxes()
      .setBackground("#212121")
      .setHorizontalAlignment("center").setVerticalAlignment("middle");

    dashboard.getRange("C1")
      .setValue("UPDATE")
      .setBackground("#212121").setFontColor("white").setFontWeight("bold")
      .setHorizontalAlignment("center").setVerticalAlignment("middle");

    dashboard.getRange("D1")
      .insertCheckboxes()
      .setBackground("#212121")
      .setHorizontalAlignment("center").setVerticalAlignment("middle");

    dashboard.getRange("E1:G1").merge()
      .setBackground("#000000").setFontColor("#00FF00").setFontWeight("bold").setFontSize(9)
      .setHorizontalAlignment("center").setVerticalAlignment("middle");

    // --- Row 2 group headers ---
    const styleGroup = (a1, label, bg) => {
      dashboard.getRange(a1).merge()
        .setValue(label)
        .setBackground(bg).setFontColor("white").setFontWeight("bold")
        .setHorizontalAlignment("center").setVerticalAlignment("middle");
    };

    dashboard.getRange("A2:AC2").clearContent();
    styleGroup("A2:A2", "IDENTITY", "#263238");
    styleGroup("B2:D2", "SIGNALING", "#0D47A1");
    styleGroup("E2:G2", "PRICE / VOLUME", "#1B5E20");
    styleGroup("H2:J2", "PERFORMANCE", "#004D40");
    styleGroup("K2:O2", "TREND", "#2E7D32");
    styleGroup("P2:T2", "MOMENTUM", "#33691E");
    styleGroup("U2:Y2", "LEVELS / RISK", "#B71C1C");
    styleGroup("Z2:Z2", "INSTITUTIONAL", "#4A148C");
    styleGroup("AA2:AB2", "NOTES", "#212121");
    styleGroup("AC2:AC2", "STATE", "#263238");
    dashboard.getRange("A2:AC2").setWrap(true);

    // --- Row 3 column headers ---
    const headers = [[
      "Ticker", "SIGNAL", "FUNDAMENTAL", "DECISION", "Price", "Change %", "Vol Trend",
      "ATH (TRUE)", "ATH Diff %", "R:R Quality", "Trend Score", "Trend State",
      "SMA 20", "SMA 50", "SMA 200",
      "RSI", "MACD Hist", "Divergence", "ADX (14)", "Stoch %K (14)",
      "Support", "Resistance", "Target (3:1)", "ATR (14)", "Bollinger %B",
      "POSITION SIZE", "TECH NOTES", "FUND NOTES"
    ]];

    dashboard.getRange(3, 1, 1, 28)
      .setValues(headers)
      .setBackground("#111111").setFontColor("white").setFontWeight("bold")
      .setHorizontalAlignment("center").setVerticalAlignment("middle")
      .setWrap(true);

    // Freeze panes
    dashboard.setFrozenRows(3);
    dashboard.setFrozenColumns(1);

    // White border for top header rows (1..3)
    dashboard.getRange("A1:AB3")
      .setBorder(true, true, true, true, true, true, "#FFFFFF", SpreadsheetApp.BorderStyle.SOLID);

    // Sentinel note
    dashboard.getRange("A1").setNote(SENTINEL);
  }
  // ==========================
  // FAST REFRESH (DATA ONLY)
  // ==========================
  dashboard.getRange(DATA_START_ROW, 1, 1000, 29).clearContent();

  // Timestamp refresh
  dashboard.getRange("E1:G1")
    .setValue(Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "MMM dd, yyyy | HH:mm:ss"));

  // Filter formula (always re-written)
  const filterFormula =
    '=IFERROR(' +
    'SORT(' +
    'FILTER({' +
    'CALCULATIONS!$A$3:$A,' +
    'CALCULATIONS!$B$3:$B,' +
    'CALCULATIONS!$C$3:$C,' +
    'CALCULATIONS!$D$3:$D,' +
    'CALCULATIONS!$E$3:$E,' +
    'CALCULATIONS!$F$3:$F,' +
    'CALCULATIONS!$G$3:$G,' +
    'CALCULATIONS!$H$3:$H,' +
    'CALCULATIONS!$I$3:$I,' +
    'CALCULATIONS!$J$3:$J,' +
    'CALCULATIONS!$K$3:$K,' +
    'CALCULATIONS!$L$3:$L,' +
    'CALCULATIONS!$M$3:$M,' +
    'CALCULATIONS!$N$3:$N,' +
    'CALCULATIONS!$O$3:$O,' +
    'CALCULATIONS!$P$3:$P,' +
    'CALCULATIONS!$Q$3:$Q,' +
    'CALCULATIONS!$R$3:$R,' +
    'CALCULATIONS!$S$3:$S,' +
    'CALCULATIONS!$T$3:$T,' +
    'CALCULATIONS!$U$3:$U,' +
    'CALCULATIONS!$V$3:$V,' +
    'CALCULATIONS!$W$3:$W,' +
    'CALCULATIONS!$X$3:$X,' +
    'CALCULATIONS!$Y$3:$Y,' +
    'CALCULATIONS!$Z$3:$Z,' +
    'CALCULATIONS!$AA$3:$AA,' +
    'CALCULATIONS!$AB$3:$AB' +
    '},' +
    'ISNUMBER(MATCH(' +
    'CALCULATIONS!$A$3:$A,' +
    'FILTER(INPUT!$A$3:$A,' +
    'INPUT!$A$3:$A<>"",' +
    '(' +
    'IF(' +
    'OR(' +
    'INPUT!$B$1="",' +
    'REGEXMATCH(UPPER(INPUT!$B$1),"(^|,\\s*)ALL(\\s*|,|$)")' +
    '),' +
    'TRUE,' +
    'REGEXMATCH(' +
    '","&UPPER(TRIM(INPUT!$B$3:$B))&"," ,' +
    '",\\s*(" & REGEXREPLACE(UPPER(TRIM(INPUT!$B$1)),"\\s*,\\s*","|") & ")\\s*,"' +
    ')' +
    ')' +
    ')' +
    '*' +
    '(' +
    'IF(' +
    'OR(' +
    'INPUT!$C$1="",' +
    'REGEXMATCH(UPPER(INPUT!$C$1),"(^|,\\s*)ALL(\\s*|,|$)")' +
    '),' +
    'TRUE,' +
    'REGEXMATCH(' +
    '","&REGEXREPLACE(UPPER(TRIM(INPUT!$C$3:$C)),"\\s+","")&"," ,' +
    '",\\s*(" & REGEXREPLACE(REGEXREPLACE(UPPER(TRIM(INPUT!$C$1)),"\\s+",""),"\\s*,\\s*","|") & ")\\s*,"' +
    ')' +
    ')' +
    ')' +
    '),0)' +
    '))' +
    ',6,FALSE' +
    '),' +
    '"No Matches Found")';

  dashboard.getRange("A4").setFormula(filterFormula);

  SpreadsheetApp.flush();

  // Apply Bloomberg formatting + heatmap ONCE
  applyDashboardBloombergFormatting_(dashboard, DATA_START_ROW);
  applyDashboardGroupMapAndColors_(dashboard);
}

function applyDashboardBloombergFormatting_(sh, DATA_START_ROW) {
  if (!sh) return;

  // ---------------------------
  // Theme (strict 3 colors)
  // ---------------------------
  const C_WHITE = "#FFFFFF";
  const C_GREEN = "#C6EFCE";
  const C_RED = "#FFC7CE";
  const C_GREY = "#E7E6E6";

  const HEADER_DARK = "#1F1F1F"; // header grey
  const HEADER_MID = "#E7E6E6"; // header grey (light)

  // Data columns in DASHBOARD (A..AB = 28); notes are now AA(27) and AB(28)
  const TOTAL_COLS = 28;

  // Column index map (Dashboard layout you use)
  // A Ticker
  // B SIGNAL
  // C FUNDAMENTAL
  // D DECISION
  // E Price
  // F Change %
  // G Vol Trend
  // H ATH
  // I ATH Diff %
  // J R:R
  // K Trend Score
  // L Trend State
  // M SMA20
  // N SMA50
  // O SMA200
  // P RSI
  // Q MACD Hist
  // R Divergence
  // S ADX
  // T StochK
  // U Support
  // V Resistance
  // W Target
  // X ATR
  // Y %B
  // Z TECH NOTES (hidden)
  // AA FUND NOTES (hidden)

  // ---------------------------
  // Helpers
  // ---------------------------
  const clamp = (n, lo, hi) => Math.max(lo, Math.min(hi, n));

  function findLastDataRow_() {
    // We determine the actual spill length by scanning column A from DATA_START_ROW down.
    // This avoids "retained colors" below when list shrinks.
    const maxScan = 2000; // bounded for performance
    const lastRowPossible = Math.max(sh.getLastRow(), DATA_START_ROW);
    const scanRows = clamp(lastRowPossible - DATA_START_ROW + 1, 1, maxScan);

    const vals = sh.getRange(DATA_START_ROW, 1, scanRows, 1).getDisplayValues().flat();
    let lastNonEmptyOffset = -1;
    for (let i = 0; i < vals.length; i++) {
      if (String(vals[i] || "").trim() !== "") lastNonEmptyOffset = i;
    }
    if (lastNonEmptyOffset === -1) return DATA_START_ROW; // no data
    return DATA_START_ROW + lastNonEmptyOffset;
  }

  function safeHideNotes_() {
    // AA = 27, AB = 28 (TECH NOTES, FUND NOTES)
    try {
      sh.hideColumns(27); // AA - TECH NOTES
      sh.hideColumns(28); // AB - FUND NOTES
    } catch (_) {
      // ignore if already hidden / protected
    }
  }

  function clearTailFormats_(lastDataRow) {
    const maxRows = sh.getMaxRows();
    const tailStart = lastDataRow + 1;
    if (tailStart <= maxRows) {
      const tailRows = maxRows - tailStart + 1;
      if (tailRows > 0) {
        sh.getRange(tailStart, 1, tailRows, TOTAL_COLS).clearFormat().clearContent();
        // NOTE: We clearContent too, to avoid ghosts in formulas spill collisions.
        // If you do NOT want tail content cleared, remove .clearContent()
      }
    }
  }

  // ---------------------------
  // Compute active window
  // ---------------------------
  safeHideNotes_();

  const lastDataRow = findLastDataRow_();
  const numRows = Math.max(1, lastDataRow - DATA_START_ROW + 1);

  const dataRange = sh.getRange(DATA_START_ROW, 1, numRows, 27); // A..AA (visible: up to POSITION SIZE)
  const borderRange = sh.getRange(DATA_START_ROW, 1, numRows, 27);

  // ---------------------------
  // Header styling (rows 1–3)
  // ---------------------------
  // Row 1: control bar (A1:AA1)
  sh.getRange(1, 1, 1, TOTAL_COLS)
    .setBackground(HEADER_DARK)
    .setFontColor("#FFFFFF")
    .setFontWeight("bold")
    .setVerticalAlignment("middle");

  // Keep your existing checkbox cells readable (B1, D1) — do not override content/checkbox
  sh.getRange("A1:D1").setHorizontalAlignment("center");
  sh.getRange("E1:G1").setHorizontalAlignment("center");

  // Row 2: group headers bar (A2:AA2)
  sh.getRange(2, 1, 1, TOTAL_COLS)
    .setBackground(HEADER_MID)
    .setFontColor("#000000")
    .setFontWeight("bold")
    .setVerticalAlignment("middle")
    .setHorizontalAlignment("center");

  // Row 3: column headers (A3:AA3)
  sh.getRange(3, 1, 1, TOTAL_COLS)
    .setBackground(HEADER_DARK)
    .setFontColor("#FFFFFF")
    .setFontWeight("bold")
    .setVerticalAlignment("middle")
    .setHorizontalAlignment("center")
    .setWrap(true);

  // ---------------------------
  // Global layout (Bloomberg dense)
  // ---------------------------
  // Column width ~ 10 chars ≈ 85 px
  for (let c = 1; c <= 27; c++) sh.setColumnWidth(c, 85);
  // Notes are hidden but keep sane widths if unhidden later
  sh.setColumnWidth(27, 420); // AA - TECH NOTES
  sh.setColumnWidth(28, 420); // AB - FUND NOTES

  // Row heights:
  // - headers: compact
  sh.setRowHeight(1, 22);
  sh.setRowHeight(2, 18);
  sh.setRowHeight(3, 22);

  // - data rows: ~ 3 lines (18 * 3 = 54)
  sh.setRowHeights(DATA_START_ROW, numRows, 54);

  // Alignment / wrap
  dataRange
    .setBackground(C_WHITE)
    .setHorizontalAlignment("left")
    .setVerticalAlignment("middle")
    .setWrap(true);

  // Clip long text in visible data area (A..Y) to keep terminal dense
  dataRange.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);

  // Borders: black for data rows
  borderRange.setBorder(true, true, true, true, true, true, "#000000", SpreadsheetApp.BorderStyle.SOLID);

  // Keep header borders clean/white (Bloomberg top bar look)
  sh.getRange(1, 1, 3, TOTAL_COLS)
    .setBorder(true, true, true, true, true, true, "#FFFFFF", SpreadsheetApp.BorderStyle.SOLID);

  // ---------------------------
  // Number formats (A..Y)
  // ---------------------------
  // Price
  sh.getRange(DATA_START_ROW, 5, numRows, 1).setNumberFormat("#,##0.00");  // E
  // Change%
  sh.getRange(DATA_START_ROW, 6, numRows, 1).setNumberFormat("0.00%");     // F
  // RVOL
  sh.getRange(DATA_START_ROW, 7, numRows, 1).setNumberFormat("0.00");      // G
  // ATH
  sh.getRange(DATA_START_ROW, 8, numRows, 1).setNumberFormat("#,##0.00");  // H
  // ATH Diff%
  sh.getRange(DATA_START_ROW, 9, numRows, 1).setNumberFormat("0.00%");     // I
  // R:R
  sh.getRange(DATA_START_ROW, 10, numRows, 1).setNumberFormat("0.00");      // J
  // SMAs
  sh.getRange(DATA_START_ROW, 13, numRows, 3).setNumberFormat("#,##0.00");  // M:N:O
  // RSI, ADX
  sh.getRange(DATA_START_ROW, 16, numRows, 1).setNumberFormat("0.0");       // P
  sh.getRange(DATA_START_ROW, 19, numRows, 1).setNumberFormat("0.0");       // S
  // MACD
  sh.getRange(DATA_START_ROW, 17, numRows, 1).setNumberFormat("0.000");     // Q
  // Stoch (0..1)
  sh.getRange(DATA_START_ROW, 20, numRows, 1).setNumberFormat("0.00%");     // T
  // Support/Res/Target/ATR
  sh.getRange(DATA_START_ROW, 21, numRows, 4).setNumberFormat("#,##0.00");  // U..X
  // %B
  sh.getRange(DATA_START_ROW, 25, numRows, 1).setNumberFormat("0.00");      // Y

  // ---------------------------
  // Clear any previous conditional formatting then apply new rules
  // ---------------------------
  const rules = [];
  const r0 = DATA_START_ROW;

  const rngCol = (col) => sh.getRange(r0, col, numRows, 1);

  const add = (formula, color, col) => {
    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied(formula)
        .setBackground(color)
        .setRanges([rngCol(col)])
        .build()
    );
  };

  // ---- SIGNAL (B) — replicate your hierarchy using existing computed SIGNAL text ----
  // Green = breakout / trend continuation / mean reversion
  // Red   = stop-out / risk-off
  // Grey  = squeeze / range / hold
  add(`=REGEXMATCH($B${r0},"Breakout|Trend Continuation|Mean Reversion")`, C_GREEN, 2);
  add(`=REGEXMATCH($B${r0},"Stop-Out|Risk-Off")`, C_RED, 2);
  add(`=REGEXMATCH($B${r0},"Volatility Squeeze|Range-Bound|Hold")`, C_GREY, 2);

  // ---- FUNDAMENTAL (C) ----
  // Green = VALUE
  // Grey  = FAIR
  // Red   = EXPENSIVE / PRICED FOR PERFECTION / ZOMBIE
  add(`=$C${r0}="VALUE"`, C_GREEN, 3);
  add(`=$C${r0}="FAIR"`, C_GREY, 3);
  add(`=REGEXMATCH($C${r0},"EXPENSIVE|PRICED FOR PERFECTION|ZOMBIE")`, C_RED, 3);

  // ---- DECISION (D) ----
  // Green = Trade Long / Accumulate / Add in Dip
  // Red   = Stop-Out / Avoid / Reduce / Take Profit
  // Grey  = Hold / Monitor / LOADING
  add(`=REGEXMATCH($D${r0},"Trade Long|Accumulate|Add in Dip")`, C_GREEN, 4);
  add(`=REGEXMATCH($D${r0},"Stop-Out|Avoid|Reduce|Take Profit")`, C_RED, 4);
  add(`=REGEXMATCH($D${r0},"Hold|Monitor|LOADING")`, C_GREY, 4);

  // ---- PRICE (E) and Change% (F) ----
  add(`=$F${r0}>0`, C_GREEN, 5);
  add(`=$F${r0}<0`, C_RED, 5);
  add(`=OR($F${r0}=0,$F${r0}="")`, C_GREY, 5);

  add(`=$F${r0}>0`, C_GREEN, 6);
  add(`=$F${r0}<0`, C_RED, 6);
  add(`=OR($F${r0}=0,$F${r0}="")`, C_GREY, 6);

  // ---- Vol Trend RVOL (G) ----
  add(`=$G${r0}>=1.5`, C_GREEN, 7);
  add(`=$G${r0}<=0.85`, C_RED, 7);
  add(`=AND($G${r0}>0.85,$G${r0}<1.5)`, C_GREY, 7);

  // ---- ATH (H) / ATH Diff % (I) ----
  // Near ATH: green; deep below ATH: red; else grey
  add(`=AND($H${r0}>0,$E${r0}>=$H${r0}*0.995)`, C_GREEN, 8);
  add(`=AND($H${r0}>0,$E${r0}<=$H${r0}*0.80)`, C_RED, 8);
  add(`=AND($H${r0}>0,$E${r0}>$H${r0}*0.80,$E${r0}<$H${r0}*0.995)`, C_GREY, 8);

  add(`=$I${r0}>=-0.05`, C_GREEN, 9);
  add(`=$I${r0}<=-0.20`, C_RED, 9);
  add(`=AND($I${r0}>-0.20,$I${r0}<-0.05)`, C_GREY, 9);

  // ---- R:R (J) ----
  add(`=$J${r0}>=3`, C_GREEN, 10);
  add(`=$J${r0}<1.5`, C_RED, 10);
  add(`=AND($J${r0}>=1.5,$J${r0}<3)`, C_GREY, 10);

  // ---- Trend Score (K) — star count ----
  add(`=LEN($K${r0})>=3`, C_GREEN, 11);
  add(`=LEN($K${r0})<=1`, C_RED, 11);
  add(`=LEN($K${r0})=2`, C_GREY, 11);

  // ---- Trend State (L) ----
  add(`=$L${r0}="BULL"`, C_GREEN, 12);
  add(`=$L${r0}="BEAR"`, C_RED, 12);
  add(`=AND($L${r0}<>"BULL",$L${r0}<>"BEAR")`, C_GREY, 12);

  // ---- SMA20/50/200 (M/N/O) vs Price ----
  add(`=AND($M${r0}>0,$E${r0}>=$M${r0})`, C_GREEN, 13);
  add(`=AND($M${r0}>0,$E${r0}<$M${r0})`, C_RED, 13);

  add(`=AND($N${r0}>0,$E${r0}>=$N${r0})`, C_GREEN, 14);
  add(`=AND($N${r0}>0,$E${r0}<$N${r0})`, C_RED, 14);

  add(`=AND($O${r0}>0,$E${r0}>=$O${r0})`, C_GREEN, 15);
  add(`=AND($O${r0}>0,$E${r0}<$O${r0})`, C_RED, 15);

  // ---- RSI (P) ----
  add(`=$P${r0}<=30`, C_GREEN, 16);
  add(`=$P${r0}>=70`, C_RED, 16);
  add(`=AND($P${r0}>30,$P${r0}<70)`, C_GREY, 16);

  // ---- MACD Hist (Q) ----
  add(`=$Q${r0}>0`, C_GREEN, 17);
  add(`=$Q${r0}<0`, C_RED, 17);
  add(`=OR($Q${r0}=0,$Q${r0}="")`, C_GREY, 17);

  // ---- Divergence (R) ----
  add(`=REGEXMATCH($R${r0},"BULL")`, C_GREEN, 18);
  add(`=REGEXMATCH($R${r0},"BEAR")`, C_RED, 18);
  add(`=OR($R${r0}="—",$R${r0}="",NOT(REGEXMATCH($R${r0},"BULL|BEAR")))`, C_GREY, 18);

  // ---- ADX (S) ----
  // Strong trend (>=25) green; low trend (<15) grey; mid grey
  add(`=$S${r0}>=25`, C_GREEN, 19);
  add(`=$S${r0}<15`, C_GREY, 19);
  add(`=AND($S${r0}>=15,$S${r0}<25)`, C_GREY, 19);

  // ---- Stoch %K (T) ----
  add(`=$T${r0}<=0.2`, C_GREEN, 20);
  add(`=$T${r0}>=0.8`, C_RED, 20);
  add(`=AND($T${r0}>0.2,$T${r0}<0.8)`, C_GREY, 20);

  // ---- Support (U) ----
  // Below support = red; within +1% above support = green; else grey
  add(`=AND($U${r0}>0,$E${r0}<$U${r0})`, C_RED, 21);
  add(`=AND($U${r0}>0,$E${r0}>=$U${r0},$E${r0}<=$U${r0}*1.01)`, C_GREEN, 21);
  add(`=AND($U${r0}>0,$E${r0}>$U${r0}*1.01)`, C_GREY, 21);

  // ---- Resistance (V) ----
  // Near/at resistance = red; far below resistance = green; else grey
  add(`=AND($V${r0}>0,$E${r0}>=$V${r0}*0.995)`, C_RED, 22);
  add(`=AND($V${r0}>0,$E${r0}<=$V${r0}*0.90)`, C_GREEN, 22);
  add(`=AND($V${r0}>0,$E${r0}>$V${r0}*0.90,$E${r0}<$V${r0}*0.995)`, C_GREY, 22);

  // ---- Target (W) ----
  // Target meaningfully above price = green; too close = red; else grey
  add(`=AND($W${r0}>0,$W${r0}>=$E${r0}*1.05)`, C_GREEN, 23);
  add(`=AND($W${r0}>0,$W${r0}<=$E${r0}*1.01)`, C_RED, 23);
  add(`=AND($W${r0}>0,$W${r0}>$E${r0}*1.01,$W${r0}<$E${r0}*1.05)`, C_GREY, 23);

  // ---- ATR (X) as % of price ----
  // Low volatility <=2% = green; high volatility >=5% = red; else grey
  add(`=IFERROR($X${r0}/$E${r0},0)<=0.02`, C_GREEN, 24);
  add(`=IFERROR($X${r0}/$E${r0},0)>=0.05`, C_RED, 24);
  add(`=AND(IFERROR($X${r0}/$E${r0},0)>0.02,IFERROR($X${r0}/$E${r0},0)<0.05)`, C_GREY, 24);

  // ---- Bollinger %B (Y) ----
  add(`=$Y${r0}<=0.2`, C_GREEN, 25);
  add(`=$Y${r0}>=0.8`, C_RED, 25);
  add(`=AND($Y${r0}>0.2,$Y${r0}<0.8)`, C_GREY, 25);

  // Apply rules
  sh.setConditionalFormatRules(rules);

  // ---------------------------
  // Hard-hide notes columns (Z, AA)
  // ---------------------------
  safeHideNotes_();

  // ---------------------------
  // Cleanup below actual data end (prevents retained row colors when shrink)
  // ---------------------------
  clearTailFormats_(lastDataRow);
}

/**
 * This colors Row 2 group bars and Row 3 headers using the same group color blocks.
 */
function applyDashboardGroupMapAndColors_(sh) {
  if (!sh) return;

  // ===== GROUP COLOR PALETTE (header-only) =====
  const COLORS = {
    SIGNAL: "#1F4FD8", // blue
    PRICE: "#0F766E", // teal
    PERF: "#374151", // slate
    TREND: "#14532D", // green
    MOM: "#7C2D12", // brown
    LEVELS: "#4C1D95", // purple
    NOTES: "#111827"  // dark
  };

  const FG = "#FFFFFF"; // white text

  // ===== GROUP → COLUMN MAP (1-indexed) =====
  const GROUPS = [
    { name: "SIGNALING", c1: 2, c2: 4, color: COLORS.SIGNAL }, // B..D
    { name: "PRICE / VOLUME", c1: 5, c2: 7, color: COLORS.PRICE }, // E..G
    { name: "PERFORMANCE", c1: 8, c2: 10, color: COLORS.PERF }, // H..J
    { name: "TREND", c1: 11, c2: 15, color: COLORS.TREND }, // K..O
    { name: "MOMENTUM", c1: 16, c2: 20, color: COLORS.MOM }, // P..T
    { name: "LEVELS / RISK", c1: 21, c2: 25, color: COLORS.LEVELS }, // U..Y
    { name: "INSTITUTIONAL", c1: 26, c2: 26, color: "#4A148C" }, // Z (Position Size)
    { name: "NOTES", c1: 27, c2: 28, color: COLORS.NOTES }  // AA..AB (Tech + Fund Notes)
  ];

  // ===== COMMON HEADER STYLE =====
  const style = (row, c1, c2, bg) => {
    sh.getRange(row, c1, 1, c2 - c1 + 1)
      .setBackground(bg)
      .setFontColor(FG)
      .setFontWeight("bold")
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle")
      .setWrap(true);
  };

  // ===== APPLY COLORS =====
  GROUPS.forEach(g => {
    // Row 2: group header bar
    style(2, g.c1, g.c2, g.color);

    // Merge + label row 2
    const r2 = sh.getRange(2, g.c1, 1, g.c2 - g.c1 + 1);
    try { r2.breakApart(); } catch (e) { }
    if (g.c1 !== g.c2) r2.merge();
    r2.setValue(g.name);

    // Row 3: column headers (same group color)
    style(3, g.c1, g.c2, g.color);
  });
}



/**
 * Call helper — keep generateDashboardSheet clean.
 * Call this AFTER you set A4 formula and flush.
 */
function applyDashboardBloombergFormattingAfterRefresh_(dashboardSheet) {
  SpreadsheetApp.flush(); // ensure FILTER() spill exists
  applyDashboardBloombergFormatting_(dashboardSheet, 4); // data starts at row 4
}

// ------------------------------------------------------------
// CHART SHEET setup engine
// ------------------------------------------------------------

function setupChartSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const input = ss.getSheetByName("INPUT");
  const calc = ss.getSheetByName("CALCULATIONS");
  if (!input || !calc) throw new Error("Missing INPUT or CALCULATIONS sheet");

  const tickers = getCleanTickers(input);
  let sh = ss.getSheetByName("CHART") || ss.insertSheet("CHART");
  sh.clear().clearFormats();
  forceExpandSheet(sh, 60);

  // ------------------------------------------------------------
  // Column sizing / alignment
  // ------------------------------------------------------------
  sh.setColumnWidth(1, 85);     // A
  sh.setColumnWidth(2, 125);    // B
  sh.setColumnWidth(3, 520);    // C Tech Notes
  sh.setColumnWidth(4, 520);    // D Fund Notes
  sh.setColumnWidth(5, 18);     // E spacer

  sh.getRange("A:A").setHorizontalAlignment("left");
  sh.getRange("B:B").setHorizontalAlignment("left").setWrap(true);

  // Dense top area
  sh.setRowHeights(1, 7, 18);

  // ------------------------------------------------------------
  // Control panel A1:B6
  // ------------------------------------------------------------
  sh.getRange("A1:B6")
    .setBackground("#000000")
    .setFontColor("#FFFF00")
    .setBorder(true, true, true, true, true, true, "#FFFFFF", SpreadsheetApp.BorderStyle.SOLID)
    .setVerticalAlignment("middle");

  // Ticker in merged A1:B1 (value lives in A1)
  sh.getRange("A1:B1").merge()
    .setValue(tickers[0] || "")
    .setFontWeight("bold")
    .setFontSize(11)
    .setHorizontalAlignment("center")
    .setFontColor("#FF80AB")
    .setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInList(tickers.length ? tickers : [""], true)
        .build()
    );

  sh.getRange("A2:A6").setValues([["YEAR"], ["MONTH"], ["DAY"], ["DATE"], ["INTERVAL"]]).setFontWeight("bold");

  const listValidation = (arr) => SpreadsheetApp.newDataValidation().requireValueInList(arr, true).build();

  // B2/B3/B4 start at 0; defaults
  sh.getRange("B2").setDataValidation(listValidation(Array.from({ length: 11 }, (_, i) => i))).setValue(1).setFontColor("#FF80AB");
  sh.getRange("B3").setDataValidation(listValidation(Array.from({ length: 13 }, (_, i) => i))).setValue(0).setFontColor("#FF80AB");
  sh.getRange("B4").setDataValidation(listValidation(Array.from({ length: 32 }, (_, i) => i))).setValue(0).setFontColor("#FF80AB");

  // Date = TODAY() minus (years+months+days)
  sh.getRange("B5").setFormula("=EDATE(TODAY(), -(12*B2+B3)) - B4").setNumberFormat("yyyy-mm-dd").setFontColor("#FF80AB");

  sh.getRange("B6")
    .setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(["DAILY", "WEEKLY"], true).build())
    .setValue("DAILY")
    .setFontWeight("bold")
    .setFontColor("#FF80AB");

  // ------------------------------------------------------------
  // Reasons: C1:C6 and D1:D6
  // CALCULATIONS: Z=TECH NOTES, AA=FUND NOTES
  // ------------------------------------------------------------
  sh.getRange("C1:C6").merge()
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP)
    .setVerticalAlignment("top")
    .setHorizontalAlignment("left")
    .setFontSize(10)
    .setFontColor("#FFD54F")
    .setBackground("#0B0B0B")
    .setBorder(true, true, true, true, true, true, "#FFFFFF", SpreadsheetApp.BorderStyle.SOLID)
    .setFormula('=IFERROR(INDEX(CALCULATIONS!$AA$3:$AA, MATCH($A$1, CALCULATIONS!$A$3:$A, 0)), "—")');

  sh.getRange("D1:D6").merge()
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP)
    .setVerticalAlignment("top")
    .setHorizontalAlignment("left")
    .setFontSize(10)
    .setFontColor("#FFD54F")
    .setBackground("#0B0B0B")
    .setBorder(true, true, true, true, true, true, "#FFFFFF", SpreadsheetApp.BorderStyle.SOLID)
    .setFormula('=IFERROR(INDEX(CALCULATIONS!$AB$3:$AB, MATCH($A$1, CALCULATIONS!$A$3:$A, 0)), "—")');

  // ------------------------------------------------------------
  // ROW 7: DECISION moved here (A7/B7) + yellow highlight
  // (Do NOT break column mapping: DECISION = CALCULATIONS column C)
  // ------------------------------------------------------------
  const t = "$A$1";
  const IDX = (colLetter, fallback) =>
    `=IFERROR(INDEX(CALCULATIONS!$${colLetter}$3:$${colLetter}, MATCH(${t}, CALCULATIONS!$A$3:$A, 0)), ${fallback})`;

  sh.getRange("A7").setValue("DECISION").setFontWeight("bold");
  sh.getRange("B7").setFormula(IDX("C", '"-"')).setFontWeight("bold");

  sh.getRange("A7:B7")
    .setBackground("#FFEB3B")
    .setFontColor("#111111")
    .setBorder(true, true, true, true, true, true, "#000000", SpreadsheetApp.BorderStyle.SOLID)
    .setVerticalAlignment("middle");

  sh.setRowHeight(7, 18);

  // ------------------------------------------------------------
  // Sidebar (starts row 8)
  // - Add borders
  // - Insert P/E and EPS under [ PERFORMANCE ]
  // - Keep all existing column mappings intact
  // ------------------------------------------------------------
  const startRow = 8;

  // Clear sidebar area (but do not touch chart data region)
  sh.getRange("A8:B200").clearContent();

  const rows = [
    ["SIGNAL", IDX("B", '"Wait"')],
    ["FUND", IDX("D", '"-"')],           // FUNDAMENTAL (CALC D)
    // DECISION removed from sidebar because moved to row 7
    ["PRICE", `=IFERROR(GOOGLEFINANCE(${t}, "price"), 0)`],
    ["CHG%", `=IFERROR(GOOGLEFINANCE(${t}, "changepct")/100, 0)`],
    ["R:R", IDX("J", "0")],
    ["", ""],

    ["[ PERFORMANCE ]", ""],
    ["VOL TREND", IDX("G", "0")],
    ["P/E", `=IFERROR(GOOGLEFINANCE(${t},"pe"), "")`],
    ["EPS", `=IFERROR(GOOGLEFINANCE(${t},"eps"), "")`],
    ["ATH", IDX("H", "0")],
    ["ATH %", IDX("I", "0")],
    ["52W HIGH", `=IFERROR(GOOGLEFINANCE(${t},"high52"), 0)`],
    ["52W LOW", `=IFERROR(GOOGLEFINANCE(${t},"low52"), 0)`],
    ["", ""],

    ["[ TREND ]", ""],
    ["SMA 20", IDX("M", "0")],
    ["SMA 50", IDX("N", "0")],
    ["SMA 200", IDX("O", "0")],
    ["RSI", IDX("P", "50")],
    ["MACD", IDX("Q", "0")],
    ["DIV", IDX("R", '"-"')],
    ["ADX", IDX("S", "0")],
    ["STO", IDX("T", "0")],
    ["", ""],

    ["[ LEVELS ]", ""],
    ["SUPPORT", IDX("U", "0")],
    ["RESISTANCE", IDX("V", "0")],
    ["TARGET", IDX("W", "0")],
    ["ATR", IDX("X", "0")],
    ["%B", IDX("Y", "0")]
  ];

  sh.getRange(startRow, 1, rows.length, 1).setValues(rows.map(r => [r[0]])).setFontWeight("bold");
  sh.getRange(startRow, 2, rows.length, 1).setFormulas(rows.map(r => [r[1]]));

  // Sidebar borders (requested)
  sh.getRange(startRow, 1, rows.length, 2)
    .setBorder(true, true, true, true, true, true, "#000000", SpreadsheetApp.BorderStyle.SOLID)
    .setVerticalAlignment("middle");

  // Style section headers
  rows.forEach((r, i) => {
    const label = String(r[0] || "");
    if (label.startsWith("[")) {
      sh.getRange(startRow + i, 1, 1, 2)
        .setBackground("#424242")
        .setFontColor("white")
        .setFontWeight("bold");
    }
  });

  sh.setRowHeights(startRow, rows.length, 18);

  // ------------------------------------------------------------
  // Number formats (robust by row numbers in this fixed sidebar)
  // ------------------------------------------------------------
  // Rows are now:
  // 8 SIGNAL
  // 9 FUND
  // 10 PRICE
  // 11 CHG%
  // 12 R:R
  // 13 blank
  // 14 [PERFORMANCE]
  // 15 VOL TREND
  // 16 P/E
  // 17 EPS
  // 18 ATH
  // 19 ATH %
  // 20 blank
  // 21 [TREND]
  // 22 SMA20
  // 23 SMA50
  // 24 SMA200
  // 25 RSI
  // 26 MACD
  // 27 DIV
  // 28 ADX
  // 29 STO
  // 30 blank
  // 31 [LEVELS]
  // 32 SUPPORT
  // 33 RESISTANCE
  // 34 TARGET
  // 35 ATR
  // 36 %B

  sh.getRange("B10").setNumberFormat("#,##0.00"); // PRICE
  sh.getRange("B11").setNumberFormat("0.00%");   // CHG%
  sh.getRange("B12").setNumberFormat("0.00");    // R:R

  sh.getRange("B15").setNumberFormat("0.00");    // VOL TREND
  sh.getRange("B16").setNumberFormat("0.00");    // P/E
  sh.getRange("B17").setNumberFormat("0.00");    // EPS
  sh.getRange("B18").setNumberFormat("#,##0.00");// ATH
  sh.getRange("B19").setNumberFormat("0.00%");   // ATH %

  sh.getRange("B22:B24").setNumberFormat("#,##0.00"); // SMA 20/50/200
  sh.getRange("B25").setNumberFormat("0.00");         // RSI
  sh.getRange("B26").setNumberFormat("0.000");        // MACD
  sh.getRange("B28").setNumberFormat("0.00");         // ADX
  sh.getRange("B29").setNumberFormat("0.00%");        // STO

  sh.getRange("B32:B35").setNumberFormat("#,##0.00"); // SUPPORT/RES/TARGET/ATR
  sh.getRange("B36").setNumberFormat("0.00");        // %B

  SpreadsheetApp.flush();

  updateDynamicChart(); // ensure chart & lines appear
}

/**
* ------------------------------------------------------------------
* updateDynamicChart() — V3_6.1.1 (Live-Stitch + Today's Data)
* ------------------------------------------------------------------
*/
function updateDynamicChart() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("CHART");
  const dataSheet = ss.getSheetByName("DATA");
  if (!sheet || !dataSheet) return;

  SpreadsheetApp.flush();

  // 1. Fetch Ticker and Settings
  const ticker = String(sheet.getRange("A1").getValue() || "").trim();
  if (!ticker) return;

  const interval = String(sheet.getRange("B6").getValue() || "DAILY").toUpperCase();
  const isWeekly = interval === "WEEKLY";

  let startDate = sheet.getRange("B5").getValue();
  if (!(startDate instanceof Date)) {
    const now = new Date();
    startDate = new Date(now.getFullYear(), now.getMonth(), now.getDate() - 14);
  }

  // 2. Fetch Sidebar Levels for Chart Lines
  const sidebar = getSidebarValuesByLabels_(sheet, ["PRICE", "SUPPORT", "RESISTANCE", "SUP", "RES"]);
  const livePrice = Number(sidebar["PRICE"]) || 0;
  const supportVal = Number(sidebar["SUPPORT"]) || Number(sidebar["SUP"]) || 0;
  const resistanceVal = Number(sidebar["RESISTANCE"]) || Number(sidebar["RES"]) || 0;

  // 3. Find ticker column in DATA
  const headers = dataSheet.getRange(2, 1, 1, dataSheet.getLastColumn()).getValues()[0];
  const colIdx = headers.indexOf(ticker);
  if (colIdx === -1) return;

  // Pull 6 cols: date, open, high, low, close, volume
  const raw = dataSheet.getRange(1, colIdx + 1, dataSheet.getLastRow(), 6).getValues();

  let master = [];
  let vols = [];
  let prices = [];

  // 4. Process Historical Data
  for (let i = 4; i < raw.length; i++) {
    const d = raw[i][0];
    const close = Number(raw[i][4]);
    const vol = Number(raw[i][5]);
    if (!d || !(d instanceof Date) || !isFinite(close) || close < 0.01) continue;
    if (d < startDate) continue;
    if (isWeekly && d.getDay() !== 5) continue;

    // SMA Calculations (Spliced for historical)
    const slice = raw.slice(Math.max(4, i - 200), i + 1).map(r => Number(r[4])).filter(n => isFinite(n) && n > 0);
    const s20 = slice.length >= 20 ? Number((slice.slice(-20).reduce((a, b) => a + b, 0) / 20).toFixed(2)) : null;
    const s50 = slice.length >= 50 ? Number((slice.slice(-50).reduce((a, b) => a + b, 0) / 50).toFixed(2)) : null;
    const s200 = slice.length >= 200 ? Number((slice.slice(-200).reduce((a, b) => a + b, 0) / 200).toFixed(2)) : null;

    const prevClose = (i > 4) ? Number(raw[i - 1][4]) : close;

    master.push([
      d, close,
      (close >= prevClose) ? vol : null,
      (close < prevClose) ? vol : null,
      s20, s50, s200,
      resistanceVal || null,
      supportVal || null
    ]);

    vols.push(vol);
    prices.push(close);
    if (s20) prices.push(s20);
    if (s50) prices.push(s50);
    if (s200) prices.push(s200);
  }

  // 5. LIVE-STITCH: Add Today's Data point if missing
  const today = new Date();
  const lastDateInMaster = master.length > 0 ? master[master.length - 1][0] : null;

  if (livePrice > 0 && (!lastDateInMaster || lastDateInMaster.toDateString() !== today.toDateString())) {
    const lastHistClose = master.length > 0 ? master[master.length - 1][1] : livePrice;

    // For live SMAs, we use the historical slices + current price
    const fullCloses = raw.map(r => Number(r[4])).filter(n => isFinite(n) && n > 0);
    fullCloses.push(livePrice);

    const liveS20 = fullCloses.length >= 20 ? Number((fullCloses.slice(-20).reduce((a, b) => a + b, 0) / 20).toFixed(2)) : null;
    const liveS50 = fullCloses.length >= 50 ? Number((fullCloses.slice(-50).reduce((a, b) => a + b, 0) / 50).toFixed(2)) : null;
    const liveS200 = fullCloses.length >= 200 ? Number((fullCloses.slice(-200).reduce((a, b) => a + b, 0) / 200).toFixed(2)) : null;

    master.push([
      today, livePrice,
      (livePrice >= lastHistClose) ? (Math.max(...vols) * 0.5) : null, // Proxy Volume for Today
      (livePrice < lastHistClose) ? (Math.max(...vols) * 0.5) : null,
      liveS20, liveS50, liveS200,
      resistanceVal || null,
      supportVal || null
    ]);
    prices.push(livePrice);
  }

  // 6. Write to Data Range (Z3:AH)
  sheet.getRange(3, 26, 2000, 9).clearContent();
  if (master.length === 0) return;

  if (supportVal > 0) prices.push(supportVal);
  if (resistanceVal > 0) prices.push(resistanceVal);
  const cleanPrices = prices.filter(p => typeof p === "number" && isFinite(p) && p > 0);
  const minP = Math.min(...cleanPrices) * 0.98;
  const maxP = Math.max(...cleanPrices) * 1.02;
  const maxVol = Math.max(...vols.filter(v => isFinite(v)), 1);

  sheet.getRange(2, 26, 1, 9).setValues([["Date", "Price", "Bull Vol", "Bear Vol", "SMA 20", "SMA 50", "SMA 200", "Resistance", "Support"]]);
  sheet.getRange(3, 26, master.length, 9).setValues(master);
  sheet.getRange(3, 26, master.length, 1).setNumberFormat("dd/MM/yy");

  // 7. Rebuild COMBO Chart
  sheet.getCharts().forEach(c => sheet.removeChart(c));
  const chart = sheet.newChart()
    .setChartType(Charts.ChartType.COMBO)
    .addRange(sheet.getRange(2, 26, master.length + 1, 9))
    .setOption("useFirstRowAsHeaders", true)
    .setOption("series", {
      0: { type: "line", color: "#1A73E8", lineWidth: 1, labelInLegend: "Price" },
      1: { type: "bars", color: "#2E7D32", targetAxisIndex: 1, labelInLegend: "Bull Vol" },
      2: { type: "bars", color: "#C62828", targetAxisIndex: 1, labelInLegend: "Bear Vol" },
      3: { type: "line", color: "#FBBC04", lineWidth: 1.5, labelInLegend: "SMA 20" },
      4: { type: "line", color: "#9C27B0", lineWidth: 1.5, labelInLegend: "SMA 50" },
      5: { type: "line", color: "#FF9800", lineWidth: 2, labelInLegend: "SMA 200" },
      6: { type: "line", color: "#B71C1C", lineDashStyle: [4, 4], labelInLegend: "Resistance" },
      7: { type: "line", color: "#0D47A1", lineDashStyle: [4, 4], labelInLegend: "Support" }
    })
    .setOption("vAxes", {
      0: { viewWindow: { min: minP, max: maxP } },
      1: { viewWindow: { min: 0, max: maxVol * 4 }, format: "short" }
    })
    .setOption("legend", { position: "top" })
    .setPosition(7, 3, 0, 0)
    .setOption("width", 1150)
    .setOption("height", 650)
    .build();

  sheet.insertChart(chart);
}
