/**
* ==============================================================================
* STABLE_MASTER_ALL_CLEAN_v3.1_KIRO_OPTIMIZED
* ==============================================================================
*/

function generateDashboardSheet() {
  const startTime = new Date();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  try {
    // Validate required sheets
    const input = ss.getSheetByName("INPUT");
    if (!input) {
      ss.toast('INPUT sheet not found.', '❌ Error', 3);
      return;
    }

    const dashboard = ss.getSheetByName("DASHBOARD") || ss.insertSheet("DASHBOARD");
    
    const DATA_START_ROW = 5;
    const SENTINEL = "DASHBOARD_LAYOUT_V3_CONTROL_ROW";
    const isInitialized = (dashboard.getRange("A1").getNote() || "").indexOf(SENTINEL) !== -1;

    // ONE-TIME LAYOUT (only if not initialized)
    if (!isInitialized) {
      setupDashboardLayout(dashboard, SENTINEL);
    }

    // FAST REFRESH (DATA ONLY)
    refreshDashboardData(dashboard, ss, DATA_START_ROW);

    const elapsed = ((new Date() - startTime) / 1000).toFixed(2);
    ss.toast(`✓ DASHBOARD refreshed in ${elapsed}s`, 'Success', 3);
    
  } catch (error) {
    ss.toast(`Failed to generate DASHBOARD: ${error.message}`, '❌ Error', 5);
    Logger.log(`Error in generateDashboardSheet: ${error.stack}`);
  }
}

function setupDashboardLayout(dashboard, SENTINEL) {
  dashboard.clear().clearFormats();

  // Get locale separator
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const SEP = (/^(en|en_)/.test(ss.getSpreadsheetLocale())) ? "," : ";";

  // Professional color scheme
  const CONTROL_BG = "#1E3A5F";        // Deep blue
  const CONTROL_LABEL = "#FFD700";     // Gold
  const CONTROL_INPUT = "#2C5282";     // Medium blue
  const SORT_BG = "#0F2942";           // Darker blue
  const NIFTY_BG = "#1A237E";          // Indigo
  const SP500_BG = "#01579B";          // Blue

  // Row 1: Country filters (A1-D1)
  dashboard.getRange("A1")
    .setValue("🇺🇸 USA")
    .setBackground(CONTROL_BG).setFontColor(CONTROL_LABEL).setFontWeight("bold")
    .setFontSize(11)
    .setHorizontalAlignment("center").setVerticalAlignment("middle");

  dashboard.getRange("B1")
    .insertCheckboxes()
    .setBackground(CONTROL_INPUT)
    .setHorizontalAlignment("center").setVerticalAlignment("middle");

  dashboard.getRange("C1")
    .setValue("🇮🇳 INDIA")
    .setBackground(CONTROL_BG).setFontColor(CONTROL_LABEL).setFontWeight("bold")
    .setFontSize(11)
    .setHorizontalAlignment("center").setVerticalAlignment("middle");

  dashboard.getRange("D1")
    .insertCheckboxes()
    .setBackground(CONTROL_INPUT)
    .setHorizontalAlignment("center").setVerticalAlignment("middle");

  // Row 1: Category filter (E1-F1)
  dashboard.getRange("E1")
    .setValue("📊 Category")
    .setBackground(CONTROL_BG).setFontColor(CONTROL_LABEL).setFontWeight("bold")
    .setFontSize(11)
    .setHorizontalAlignment("center").setVerticalAlignment("middle");

  dashboard.getRange("F1")
    .setValue("ALL")
    .setBackground(CONTROL_INPUT).setFontColor("white").setFontWeight("bold")
    .setFontSize(10)
    .setHorizontalAlignment("center").setVerticalAlignment("middle");

  // Row 1: TRADE mode toggle (G1-H1)
  dashboard.getRange("G1")
    .setValue("⚡ INVEST")
    .setBackground(CONTROL_BG).setFontColor(CONTROL_LABEL).setFontWeight("bold")
    .setFontSize(11)
    .setHorizontalAlignment("center").setVerticalAlignment("middle");

  dashboard.getRange("H1")
    .insertCheckboxes()
    .setBackground(CONTROL_INPUT)
    .setHorizontalAlignment("center").setVerticalAlignment("middle");

  // Row 1: Dashboard refresh (I1-J1)
  dashboard.getRange("I1")
    .setValue("🔄 Refresh")
    .setBackground(CONTROL_BG).setFontColor(CONTROL_LABEL).setFontWeight("bold")
    .setFontSize(11)
    .setHorizontalAlignment("center").setVerticalAlignment("middle");

  dashboard.getRange("J1")
    .insertCheckboxes()
    .setBackground(CONTROL_INPUT)
    .setHorizontalAlignment("center").setVerticalAlignment("middle");

  // Row 1: Calculations refresh (K1-L1)
  dashboard.getRange("K1")
    .setValue("🧮 CALC")
    .setBackground(CONTROL_BG).setFontColor(CONTROL_LABEL).setFontWeight("bold")
    .setFontSize(11)
    .setHorizontalAlignment("center").setVerticalAlignment("middle");

  dashboard.getRange("L1")
    .insertCheckboxes()
    .setBackground(CONTROL_INPUT)
    .setHorizontalAlignment("center").setVerticalAlignment("middle");

  // Row 1: Data rebuild (M1-N1)
  dashboard.getRange("M1")
    .setValue("💾 DATA")
    .setBackground(CONTROL_BG).setFontColor(CONTROL_LABEL).setFontWeight("bold")
    .setFontSize(11)
    .setHorizontalAlignment("center").setVerticalAlignment("middle");

  dashboard.getRange("N1")
    .insertCheckboxes()
    .setBackground(CONTROL_INPUT)
    .setHorizontalAlignment("center").setVerticalAlignment("middle");

  // Row 1: Alert mode (O1-P1)
  dashboard.getRange("O1")
    .setValue("🔔 ALERT")
    .setBackground(CONTROL_BG).setFontColor(CONTROL_LABEL).setFontWeight("bold")
    .setFontSize(11)
    .setHorizontalAlignment("center").setVerticalAlignment("middle");

  dashboard.getRange("P1")
    .setValue("All")
    .setBackground(CONTROL_INPUT).setFontColor("white").setFontWeight("bold")
    .setFontSize(10)
    .setHorizontalAlignment("center").setVerticalAlignment("middle");

  // Add borders to Row 1
  dashboard.getRange("A1:P1")
    .setBorder(true, true, true, true, true, true, "#FFD700", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  // Row 2: Sort controls (A2-B2)
  dashboard.getRange("A2")
    .setValue("⬇️ Sort By")
    .setBackground(SORT_BG).setFontColor(CONTROL_LABEL).setFontWeight("bold")
    .setFontSize(10)
    .setHorizontalAlignment("center").setVerticalAlignment("middle");

  dashboard.getRange("B2")
    .setValue("Change %")
    .setBackground(CONTROL_INPUT).setFontColor("white").setFontWeight("bold")
    .setFontSize(10)
    .setHorizontalAlignment("center").setVerticalAlignment("middle");

  // Row 2: Market indices (C2-H2)
  dashboard.getRange("C2")
    .setValue("NIFTY 50")
    .setBackground(NIFTY_BG).setFontColor("white").setFontWeight("bold")
    .setFontSize(10)
    .setHorizontalAlignment("center").setVerticalAlignment("middle");

  dashboard.getRange("D2")
    .setFormula(`=IFERROR(GOOGLEFINANCE("INDEXNSE:NIFTY_50"${SEP}"price")${SEP}0)`)
    .setBackground(NIFTY_BG).setFontColor("white").setFontWeight("bold")
    .setFontSize(10)
    .setHorizontalAlignment("center").setVerticalAlignment("middle")
    .setNumberFormat("#,##0.00");

  dashboard.getRange("E2")
    .setFormula(`=IFERROR(GOOGLEFINANCE("INDEXNSE:NIFTY_50"${SEP}"changepct")/100${SEP}0)`)
    .setBackground(NIFTY_BG).setFontColor("white").setFontWeight("bold")
    .setFontSize(10)
    .setHorizontalAlignment("center").setVerticalAlignment("middle")
    .setNumberFormat("0.00%");

  dashboard.getRange("F2")
    .setValue("S&P 500")
    .setBackground(SP500_BG).setFontColor("white").setFontWeight("bold")
    .setFontSize(10)
    .setHorizontalAlignment("center").setVerticalAlignment("middle");

  dashboard.getRange("G2")
    .setFormula(`=IFERROR(GOOGLEFINANCE("INDEXSP:.INX"${SEP}"price")${SEP}0)`)
    .setBackground(SP500_BG).setFontColor("white").setFontWeight("bold")
    .setFontSize(10)
    .setHorizontalAlignment("center").setVerticalAlignment("middle")
    .setNumberFormat("#,##0.00");

  dashboard.getRange("H2")
    .setFormula(`=IFERROR(GOOGLEFINANCE("INDEXSP:.INX"${SEP}"changepct")/100${SEP}0)`)
    .setBackground(SP500_BG).setFontColor("white").setFontWeight("bold")
    .setFontSize(10)
    .setHorizontalAlignment("center").setVerticalAlignment("middle")
    .setNumberFormat("0.00%");

  // Add borders to Row 2
  dashboard.getRange("A2:H2")
    .setBorder(true, true, true, true, true, true, "#4A90E2", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  // Set row heights for better appearance
  dashboard.setRowHeight(1, 28);
  dashboard.setRowHeight(2, 26);

  // Row 3: Group headers (moved from old Row 2)
  const styleGroup = (a1, label, bg) => {
    dashboard.getRange(a1).merge()
      .setValue(label)
      .setBackground(bg).setFontColor("white").setFontWeight("bold")
      .setHorizontalAlignment("center").setVerticalAlignment("middle");
  };

  // Clear any existing merges in row 3 to avoid merge conflicts
  try {
    dashboard.getRange("A3:AG3").breakApart();
  } catch (e) {
    // Ignore if no merges exist
  }

  dashboard.getRange("A3:AG3").clearContent();
  styleGroup("A3:A3", "IDENTITY", "#37474F");        // Dark Blue-Grey (A)
  styleGroup("B3:F3", "SIGNALING", "#1565C0");       // Blue (B-F: MARKET RATING, DECISION, SIGNAL, PATTERNS, CONSENSUS PRICE)
  styleGroup("G3:I3", "PRICE / VOLUME", "#D84315");  // Deep Orange (G-I: Price, Change%, Vol Trend)
  styleGroup("J3:M3", "PERFORMANCE", "#1976D2");     // Medium Blue (J-M: ATH TRUE, ATH Diff%, ATH ZONE, FUNDAMENTAL)
  styleGroup("N3:Q3", "TREND", "#00838F");           // Cyan (N-Q: Trend State, SMA 20/50/200)
  styleGroup("R3:V3", "MOMENTUM", "#F57C00");        // Orange (R-V: RSI, MACD, Div, ADX, Stoch)
  styleGroup("W3:Z3", "VOLATILITY", "#C62828");      // Red (W-Z: VOL REGIME, BBP SIGNAL, ATR, Bollinger %B)
  styleGroup("AA3:AG3", "TARGET", "#AD1457");        // Pink (AA-AG: All target-related)
  dashboard.getRange("A3:AG3").setWrap(true);

  // Row 4: Column headers (moved from old Row 3) - 33 columns A-AG
  const headers = [[
    "Ticker",           // A
    "MARKET RATING",    // B (NEW)
    "DECISION",         // C
    "SIGNAL",           // D
    "PATTERNS",         // E
    "CONSENSUS PRICE",  // F (NEW)
    "Price",            // G
    "Change %",         // H
    "Vol Trend",        // I
    "ATH (TRUE)",       // J
    "ATH Diff %",       // K
    "ATH ZONE",         // L
    "FUNDAMENTAL",      // M
    "Trend State",      // N
    "SMA 20",           // O
    "SMA 50",           // P
    "SMA 200",          // Q
    "RSI",              // R
    "MACD Hist",        // S
    "Divergence",       // T
    "ADX (14)",         // U
    "Stoch %K (14)",    // V
    "VOL REGIME",       // W
    "BBP SIGNAL",       // X
    "ATR (14)",         // Y
    "Bollinger %B",     // Z
    "Target (3:1)",     // AA
    "R:R Quality",      // AB
    "Support",          // AC
    "Resistance",       // AD
    "ATR STOP",         // AE
    "ATR TARGET",       // AF
    "POSITION SIZE"     // AG (33 columns total)
  ]];

  dashboard.getRange(4, 1, 1, 33)
    .setValues(headers)
    .setBackground("#0D0D0D").setFontColor("#FFD700").setFontWeight("bold")
    .setHorizontalAlignment("center").setVerticalAlignment("middle")
    .setWrap(true);

  // Freeze panes
  dashboard.setFrozenRows(4);
  dashboard.setFrozenColumns(1);

  // Borders for rows 3-4 only (preserve row 1-2 borders from above)
  dashboard.getRange("A3:AG4")
    .setBorder(true, true, true, true, true, true, "#FFFFFF", SpreadsheetApp.BorderStyle.SOLID);

  // Sentinel note
  dashboard.getRange("A1").setNote(SENTINEL);
  
  // Setup dropdowns for control row
  setupControlRowDropdowns();
}

function refreshDashboardData(dashboard, ss, DATA_START_ROW, preserveCheckboxes = false) {
  // Clear existing data (33 columns A-AG)
  dashboard.getRange(DATA_START_ROW, 1, 1000, 33).clearContent();

  // Only reset checkboxes on initial setup (not when filters change)
  if (!preserveCheckboxes) {
    dashboard.getRange("B1").setValue(false);
    dashboard.getRange("D1").setValue(false);
    dashboard.getRange("H1").setValue(false);
    dashboard.getRange("J1").setValue(false);
    dashboard.getRange("L1").setValue(false);
    dashboard.getRange("N1").setValue(false);
  }

  // Get sort column from B2 (default to "Change %")
  const sortColumnName = dashboard.getRange("B2").getValue() || "Change %";
  
  // Map column names to their positions in the DASHBOARD (1-based)
  const columnMap = {
    "Ticker": 1,
    "MARKET RATING": 2,
    "DECISION": 3,
    "SIGNAL": 4,
    "PATTERNS": 5,
    "CONSENSUS PRICE": 6,
    "Price": 7,
    "Change %": 8,
    "Vol Trend": 9,
    "ATH (TRUE)": 10,
    "ATH Diff %": 11,
    "ATH ZONE": 12,
    "FUNDAMENTAL": 13,
    "Trend State": 14,
    "SMA 20": 15,
    "SMA 50": 16,
    "SMA 200": 17,
    "RSI": 18,
    "MACD Hist": 19,
    "Divergence": 20,
    "ADX (14)": 21,
    "Stoch %K (14)": 22,
    "VOL REGIME": 23,
    "BBP SIGNAL": 24,
    "ATR (14)": 25,
    "Bollinger %B": 26,
    "Target (3:1)": 27,
    "R:R Quality": 28,
    "Support": 29,
    "Resistance": 30,
    "ATR STOP": 31,
    "ATR TARGET": 32,
    "POSITION SIZE": 33
  };
  
  const sortColumnIndex = columnMap[sortColumnName] || 8; // Default to column 8 (Change %)

  // Filter formula - pulls columns from CALCULATIONS in CORRECT ORDER (matches CALCULATIONS exactly)
  const filterFormula =
    '=IFERROR(' +
    'SORT(' +
    'FILTER({' +
    'CALCULATIONS!$A$3:$A,' +   // A: Ticker
    'CALCULATIONS!$B$3:$B,' +   // B: MARKET RATING (NEW)
    'CALCULATIONS!$C$3:$C,' +   // C: DECISION (shifted from B)
    'CALCULATIONS!$D$3:$D,' +   // D: SIGNAL (shifted from C)
    'CALCULATIONS!$E$3:$E,' +   // E: PATTERNS (shifted from D)
    'CALCULATIONS!$F$3:$F,' +   // F: CONSENSUS PRICE (NEW)
    'CALCULATIONS!$G$3:$G,' +   // G: Price (shifted from E)
    'CALCULATIONS!$H$3:$H,' +   // H: Change % (shifted from F)
    'CALCULATIONS!$I$3:$I,' +   // I: Vol Trend (shifted from G)
    'CALCULATIONS!$J$3:$J,' +   // J: ATH (TRUE) (shifted from H)
    'CALCULATIONS!$K$3:$K,' +   // K: ATH Diff % (shifted from I)
    'CALCULATIONS!$L$3:$L,' +   // L: ATH ZONE (shifted from J)
    'CALCULATIONS!$M$3:$M,' +   // M: FUNDAMENTAL (shifted from K)
    'CALCULATIONS!$N$3:$N,' +   // N: Trend State (shifted from L)
    'CALCULATIONS!$O$3:$O,' +   // O: SMA 20 (shifted from M)
    'CALCULATIONS!$P$3:$P,' +   // P: SMA 50 (shifted from N)
    'CALCULATIONS!$Q$3:$Q,' +   // Q: SMA 200 (shifted from O)
    'CALCULATIONS!$R$3:$R,' +   // R: RSI (shifted from P)
    'CALCULATIONS!$S$3:$S,' +   // S: MACD Hist (shifted from Q)
    'CALCULATIONS!$T$3:$T,' +   // T: Divergence (shifted from R)
    'CALCULATIONS!$U$3:$U,' +   // U: ADX (14) (shifted from S)
    'CALCULATIONS!$V$3:$V,' +   // V: Stoch %K (14) (shifted from T)
    'CALCULATIONS!$W$3:$W,' +   // W: VOL REGIME (shifted from U)
    'CALCULATIONS!$X$3:$X,' +   // X: BBP SIGNAL (shifted from V)
    'CALCULATIONS!$Y$3:$Y,' +   // Y: ATR (14) (shifted from W)
    'CALCULATIONS!$Z$3:$Z,' +   // Z: Bollinger %B (shifted from X)
    'CALCULATIONS!$AA$3:$AA,' + // AA: Target (3:1) (shifted from Y)
    'CALCULATIONS!$AB$3:$AB,' + // AB: R:R Quality (shifted from Z)
    'CALCULATIONS!$AC$3:$AC,' + // AC: Support (shifted from AA)
    'CALCULATIONS!$AD$3:$AD,' + // AD: Resistance (shifted from AB)
    'CALCULATIONS!$AE$3:$AE,' + // AE: ATR STOP (shifted from AC)
    'CALCULATIONS!$AF$3:$AF,' + // AF: ATR TARGET (shifted from AD)
    'CALCULATIONS!$AG$3:$AG' +  // AG: POSITION SIZE (shifted from AE)
    // Note: LAST STATE (AH) is not included in DASHBOARD
    '},' +
    'ISNUMBER(MATCH(' +
    'CALCULATIONS!$A$3:$A,' +
    'FILTER(INPUT!$A$3:$A,' +
    'INPUT!$A$3:$A<>"",' +
    '(' +
    'IF(' +
    'AND(DASHBOARD!$B$1=TRUE,DASHBOARD!$D$1=TRUE),' +  // Both USA and INDIA checked
    'TRUE,' +  // Show all countries
    'IF(' +
    'DASHBOARD!$B$1=TRUE,' +  // Only USA checked
    'UPPER(TRIM(INPUT!$B$3:$B))="USA",' +
    'IF(' +
    'DASHBOARD!$D$1=TRUE,' +  // Only INDIA checked
    'UPPER(TRIM(INPUT!$B$3:$B))="INDIA",' +
    'TRUE' +  // Neither checked, show all
    ')' +
    ')' +
    ')' +
    ')' +
    '*' +
    '(' +
    'IF(' +
    'OR(' +
    'DASHBOARD!$F$1="",' +
    'REGEXMATCH(UPPER(DASHBOARD!$F$1),"(^|,\\s*)ALL(\\s*|,|$)")' +
    '),' +
    'TRUE,' +
    'REGEXMATCH(' +
    '","&REGEXREPLACE(UPPER(TRIM(INPUT!$C$3:$C)),"\\s+","")&"," ,' +
    '",\\s*(" & REGEXREPLACE(REGEXREPLACE(UPPER(TRIM(DASHBOARD!$F$1)),"\\s+",""),"\\s*,\\s*","|") & ")\\s*,"' +
    ')' +
    ')' +
    ')' +
    '),0)' +
    '))' +
    ',' + sortColumnIndex + ',FALSE' +  // Dynamic sort column, descending (FALSE = DESC)
    '),' +
    '"No Matches Found")';

  dashboard.getRange("A5").setFormula(filterFormula);
  SpreadsheetApp.flush();

  // ONLY apply data cell formatting (not headers, not control rows)
  applyDataCellFormattingOnly_(dashboard, DATA_START_ROW);
}

/**
 * Apply ONLY data cell formatting (no headers, no control rows)
 * Lightweight version for filter/sort changes
 */
function applyDataCellFormattingOnly_(sh, DATA_START_ROW) {
  if (!sh) return;

  const C_BLUE = "#E3F2FD";   // Light blue (default background)
  const C_GREEN = "#C8E6C9";  // Light green (positive)
  const C_RED = "#FFCDD2";    // Light red (negative)
  const TOTAL_COLS = 33;

  const clamp = (n, lo, hi) => Math.max(lo, Math.min(hi, n));

  function findLastDataRow_() {
    const maxScan = 2000;
    const lastRowPossible = Math.max(sh.getLastRow(), DATA_START_ROW);
    const scanRows = clamp(lastRowPossible - DATA_START_ROW + 1, 1, maxScan);
    const vals = sh.getRange(DATA_START_ROW, 1, scanRows, 1).getDisplayValues().flat();
    let lastNonEmptyOffset = -1;
    for (let i = 0; i < vals.length; i++) {
      if (String(vals[i] || "").trim() !== "") lastNonEmptyOffset = i;
    }
    return lastNonEmptyOffset === -1 ? DATA_START_ROW : DATA_START_ROW + lastNonEmptyOffset;
  }

  const lastDataRow = findLastDataRow_();
  const numRows = Math.max(1, lastDataRow - DATA_START_ROW + 1);
  const dataRange = sh.getRange(DATA_START_ROW, 1, numRows, 33);

  // Data range styling - LEFT ALIGNED with borders and BLACK text
  dataRange
    .setBackground(C_BLUE)  // Light blue default background
    .setHorizontalAlignment("left")  // LEFT ALIGN all data
    .setVerticalAlignment("middle")
    .setFontColor("#000000")  // BLACK text for light backgrounds
    .setFontSize(10)
    .setWrap(true)
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);

  // Add borders to all data cells
  dataRange.setBorder(true, true, true, true, true, true, "#000000", SpreadsheetApp.BorderStyle.SOLID);

  // Number formats (updated for 33 columns A-AG)
  sh.getRange(DATA_START_ROW, 6, numRows, 1).setNumberFormat("#,##0.00");  // F: CONSENSUS PRICE
  sh.getRange(DATA_START_ROW, 7, numRows, 1).setNumberFormat("#,##0.00");  // G: Price
  sh.getRange(DATA_START_ROW, 8, numRows, 1).setNumberFormat("0.00%");     // H: Change%
  sh.getRange(DATA_START_ROW, 9, numRows, 1).setNumberFormat("0.00");      // I: Vol Trend (RVOL)
  sh.getRange(DATA_START_ROW, 10, numRows, 1).setNumberFormat("#,##0.00"); // J: ATH (TRUE)
  sh.getRange(DATA_START_ROW, 11, numRows, 1).setNumberFormat("0.00%");    // K: ATH Diff%
  sh.getRange(DATA_START_ROW, 12, numRows, 1).setNumberFormat("@");        // L: ATH ZONE
  sh.getRange(DATA_START_ROW, 13, numRows, 1).setNumberFormat("@");        // M: FUNDAMENTAL
  sh.getRange(DATA_START_ROW, 15, numRows, 3).setNumberFormat("#,##0.00"); // O-Q: SMAs
  sh.getRange(DATA_START_ROW, 18, numRows, 1).setNumberFormat("0.0");      // R: RSI
  sh.getRange(DATA_START_ROW, 19, numRows, 1).setNumberFormat("0.000");    // S: MACD
  sh.getRange(DATA_START_ROW, 21, numRows, 1).setNumberFormat("0.0");      // U: ADX
  sh.getRange(DATA_START_ROW, 22, numRows, 1).setNumberFormat("0.00%");    // V: Stoch
  sh.getRange(DATA_START_ROW, 23, numRows, 1).setNumberFormat("@");        // W: VOL REGIME
  sh.getRange(DATA_START_ROW, 24, numRows, 1).setNumberFormat("@");        // X: BBP SIGNAL
  sh.getRange(DATA_START_ROW, 25, numRows, 1).setNumberFormat("#,##0.00"); // Y: ATR
  sh.getRange(DATA_START_ROW, 26, numRows, 1).setNumberFormat("0.00");     // Z: Bollinger %B
  sh.getRange(DATA_START_ROW, 27, numRows, 1).setNumberFormat("#,##0.00"); // AA: Target (3:1)
  sh.getRange(DATA_START_ROW, 28, numRows, 1).setNumberFormat("0.00");     // AB: R:R Quality
  sh.getRange(DATA_START_ROW, 29, numRows, 1).setNumberFormat("#,##0.00"); // AC: Support
  sh.getRange(DATA_START_ROW, 30, numRows, 1).setNumberFormat("#,##0.00"); // AD: Resistance
  sh.getRange(DATA_START_ROW, 31, numRows, 1).setNumberFormat("#,##0.00"); // AE: ATR STOP
  sh.getRange(DATA_START_ROW, 32, numRows, 1).setNumberFormat("#,##0.00"); // AF: ATR TARGET
  sh.getRange(DATA_START_ROW, 33, numRows, 1).setNumberFormat("@");        // AG: POSITION SIZE

  // Apply conditional formatting rules
  applyConditionalFormatting(sh, DATA_START_ROW, numRows, C_GREEN, C_RED, C_BLUE);
}

function applyDashboardBloombergFormatting_(sh, DATA_START_ROW) {
  if (!sh) return;

  const C_BLUE = "#E3F2FD";   // Light blue (default background)
  const C_GREEN = "#C8E6C9";  // Light green (positive)
  const C_RED = "#FFCDD2";    // Light red (negative)
  const HEADER_DARK = "#1F1F1F";
  const TOTAL_COLS = 33; // Updated to 33 columns (A-AG, includes MARKET RATING and CONSENSUS PRICE)

  const clamp = (n, lo, hi) => Math.max(lo, Math.min(hi, n));

  function findLastDataRow_() {
    const maxScan = 2000;
    const lastRowPossible = Math.max(sh.getLastRow(), DATA_START_ROW);
    const scanRows = clamp(lastRowPossible - DATA_START_ROW + 1, 1, maxScan);
    const vals = sh.getRange(DATA_START_ROW, 1, scanRows, 1).getDisplayValues().flat();
    let lastNonEmptyOffset = -1;
    for (let i = 0; i < vals.length; i++) {
      if (String(vals[i] || "").trim() !== "") lastNonEmptyOffset = i;
    }
    return lastNonEmptyOffset === -1 ? DATA_START_ROW : DATA_START_ROW + lastNonEmptyOffset;
  }

  function safeHideNotes_() {
    // No columns to hide - all 32 columns are visible
  }

  function clearTailFormats_(lastDataRow) {
    const maxRows = sh.getMaxRows();
    const tailStart = lastDataRow + 1;
    if (tailStart <= maxRows) {
      const tailRows = maxRows - tailStart + 1;
      if (tailRows > 0) {
        sh.getRange(tailStart, 1, tailRows, TOTAL_COLS).clearFormat().clearContent();
      }
    }
  }

  safeHideNotes_();
  const lastDataRow = findLastDataRow_();
  const numRows = Math.max(1, lastDataRow - DATA_START_ROW + 1);
  const dataRange = sh.getRange(DATA_START_ROW, 1, numRows, 33); // Updated to 33 columns

  // Header styling
  // Row 1: Control row (all controls A1-P1) - preserve existing styling from setupDashboardLayout
  // Row 2: Sort controls and market indices - preserve existing styling from setupDashboardLayout

  // Row 3: Group headers
  sh.getRange(3, 1, 1, TOTAL_COLS)
    .setBackground("#E7E6E6")
    .setFontColor("#000000")
    .setFontWeight("bold")
    .setVerticalAlignment("middle")
    .setHorizontalAlignment("center");

  // Row 4: Column headers
  sh.getRange(4, 1, 1, TOTAL_COLS)
    .setBackground(HEADER_DARK)
    .setFontColor("#FFFFFF")
    .setFontWeight("bold")
    .setVerticalAlignment("middle")
    .setHorizontalAlignment("center")
    .setWrap(true);

  // Column widths (33 columns)
  for (let c = 1; c <= 33; c++) sh.setColumnWidth(c, 85);

  // Row heights - preserve control row heights from setupDashboardLayout
  // Row 1 and 2 heights are already set in setupDashboardLayout (28px and 26px)
  sh.setRowHeight(3, 18);
  sh.setRowHeight(4, 22);
  sh.setRowHeights(DATA_START_ROW, numRows, 54);

  // Data range styling - LEFT ALIGNED with borders and BLACK text
  dataRange
    .setBackground(C_BLUE)  // Light blue default background
    .setHorizontalAlignment("left")  // LEFT ALIGN all data
    .setVerticalAlignment("middle")
    .setFontColor("#000000")  // BLACK text for light backgrounds
    .setFontSize(10)
    .setWrap(true)
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);

  // Add borders to all data cells
  dataRange.setBorder(true, true, true, true, true, true, "#000000", SpreadsheetApp.BorderStyle.SOLID);

  // Borders for rows 3-4 only (preserve row 1-2 custom borders)
  sh.getRange(3, 1, 2, TOTAL_COLS)
    .setBorder(true, true, true, true, true, true, "#FFFFFF", SpreadsheetApp.BorderStyle.SOLID);

  // Number formats (updated for 33 columns A-AG)
  sh.getRange(DATA_START_ROW, 6, numRows, 1).setNumberFormat("#,##0.00");  // F: CONSENSUS PRICE
  sh.getRange(DATA_START_ROW, 7, numRows, 1).setNumberFormat("#,##0.00");  // G: Price
  sh.getRange(DATA_START_ROW, 8, numRows, 1).setNumberFormat("0.00%");     // H: Change%
  sh.getRange(DATA_START_ROW, 9, numRows, 1).setNumberFormat("0.00");      // I: Vol Trend (RVOL)
  sh.getRange(DATA_START_ROW, 10, numRows, 1).setNumberFormat("#,##0.00"); // J: ATH (TRUE)
  sh.getRange(DATA_START_ROW, 11, numRows, 1).setNumberFormat("0.00%");    // K: ATH Diff%
  sh.getRange(DATA_START_ROW, 12, numRows, 1).setNumberFormat("@");        // L: ATH ZONE
  sh.getRange(DATA_START_ROW, 13, numRows, 1).setNumberFormat("@");        // M: FUNDAMENTAL
  sh.getRange(DATA_START_ROW, 15, numRows, 3).setNumberFormat("#,##0.00"); // O-Q: SMAs
  sh.getRange(DATA_START_ROW, 18, numRows, 1).setNumberFormat("0.0");      // R: RSI
  sh.getRange(DATA_START_ROW, 19, numRows, 1).setNumberFormat("0.000");    // S: MACD
  sh.getRange(DATA_START_ROW, 21, numRows, 1).setNumberFormat("0.0");      // U: ADX
  sh.getRange(DATA_START_ROW, 22, numRows, 1).setNumberFormat("0.00%");    // V: Stoch
  sh.getRange(DATA_START_ROW, 23, numRows, 1).setNumberFormat("@");        // W: VOL REGIME
  sh.getRange(DATA_START_ROW, 24, numRows, 1).setNumberFormat("@");        // X: BBP SIGNAL
  sh.getRange(DATA_START_ROW, 25, numRows, 1).setNumberFormat("#,##0.00"); // Y: ATR
  sh.getRange(DATA_START_ROW, 26, numRows, 1).setNumberFormat("0.00");     // Z: Bollinger %B
  sh.getRange(DATA_START_ROW, 27, numRows, 1).setNumberFormat("#,##0.00"); // AA: Target (3:1)
  sh.getRange(DATA_START_ROW, 28, numRows, 1).setNumberFormat("0.00");     // AB: R:R Quality
  sh.getRange(DATA_START_ROW, 29, numRows, 1).setNumberFormat("#,##0.00"); // AC: Support
  sh.getRange(DATA_START_ROW, 30, numRows, 1).setNumberFormat("#,##0.00"); // AD: Resistance
  sh.getRange(DATA_START_ROW, 31, numRows, 1).setNumberFormat("#,##0.00"); // AE: ATR STOP
  sh.getRange(DATA_START_ROW, 32, numRows, 1).setNumberFormat("#,##0.00"); // AF: ATR TARGET
  sh.getRange(DATA_START_ROW, 33, numRows, 1).setNumberFormat("@");        // AG: POSITION SIZE

  // Apply conditional formatting rules
  applyConditionalFormatting(sh, DATA_START_ROW, numRows, C_GREEN, C_RED, C_BLUE);

  safeHideNotes_();
  clearTailFormats_(lastDataRow);
}

function applyConditionalFormatting(sh, r0, numRows, C_GREEN, C_RED, C_BLUE) {
  const rules = [];
  const rngCol = (col) => sh.getRange(r0, col, numRows, 1);
  const add = (formula, color, col) => {
    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied(formula)
        .setBackground(color)
        .setFontColor("#000000")  // BLACK text for light backgrounds
        .setRanges([rngCol(col)])
        .build()
    );
  };

  // MARKET RATING (B) - Green for buy ratings, Red for sell ratings
  add(`=REGEXMATCH($B${r0},"(?i)(BUY|STRONG BUY|OUTPERFORM|OVERWEIGHT)")`, C_GREEN, 2);
  add(`=REGEXMATCH($B${r0},"(?i)(SELL|STRONG SELL|UNDERPERFORM|UNDERWEIGHT)")`, C_RED, 2);
  
  // DECISION (C) - Green for buy signals, Red for sell signals
  add(`=REGEXMATCH($C${r0},"STRONG BUY|BUY|ADD|STRONG TRADE|TRADE LONG|Accumulate|Add in Dip")`, C_GREEN, 3);
  add(`=REGEXMATCH($C${r0},"EXIT|AVOID|STOP OUT|Stop-Out|Risk-Off|Take Profit|TRIM")`, C_RED, 3);

  // SIGNAL (D) - Green for bullish, Red for bearish
  add(`=REGEXMATCH($D${r0},"STRONG BUY|ATH BREAKOUT|VOLATILITY BREAKOUT|BUY|ACCUMULATE|BREAKOUT|MOMENTUM|UPTREND|BULLISH|OVERSOLD")`, C_GREEN, 4);
  add(`=REGEXMATCH($D${r0},"STOP OUT|RISK OFF")`, C_RED, 4);

  // PATTERNS (E) - Green for bullish patterns, Red for bearish
  add(`=REGEXMATCH($E${r0},"ASC_TRI|BRKOUT|DBL_BTM|INV_H&S|CUP_HDL")`, C_GREEN, 5);
  add(`=REGEXMATCH($E${r0},"DESC_TRI|H&S|DBL_TOP")`, C_RED, 5);

  // CONSENSUS PRICE (F) - Red if below Price, Green if >15% above Price
  add(`=AND($F${r0}>0,$G${r0}>0,$F${r0}<$G${r0})`, C_RED, 6);
  add(`=AND($F${r0}>0,$G${r0}>0,$F${r0}>=$G${r0}*1.15)`, C_GREEN, 6);

  // PRICE (G) and Change% (H) - Green for positive, Red for negative
  add(`=$H${r0}>0`, C_GREEN, 7);
  add(`=$H${r0}<0`, C_RED, 7);
  add(`=$H${r0}>0`, C_GREEN, 8);
  add(`=$H${r0}<0`, C_RED, 8);

  // Vol Trend RVOL (I) - Green for high volume, Red for low
  add(`=$I${r0}>=1.5`, C_GREEN, 9);
  add(`=$I${r0}<=0.85`, C_RED, 9);

  // ATH (TRUE) (J) - Green near ATH, Red far from ATH
  add(`=AND($J${r0}>0,$G${r0}>=$J${r0}*0.995)`, C_GREEN, 10);
  add(`=AND($J${r0}>0,$G${r0}<=$J${r0}*0.80)`, C_RED, 10);

  // ATH Diff % (K) - Green near ATH, Red far from ATH
  add(`=$K${r0}>=-0.05`, C_GREEN, 11);
  add(`=$K${r0}<=-0.20`, C_RED, 11);

  // ATH ZONE (L) - Green at/near ATH, Red in correction
  add(`=REGEXMATCH($L${r0},"AT ATH|NEAR ATH")`, C_GREEN, 12);
  add(`=REGEXMATCH($L${r0},"DEEP VALUE|CORRECTION")`, C_RED, 12);

  // FUNDAMENTAL (M) - Green for value, Red for expensive
  add(`=$M${r0}="VALUE"`, C_GREEN, 13);
  add(`=REGEXMATCH($M${r0},"EXPENSIVE|PRICED FOR PERFECTION|ZOMBIE")`, C_RED, 13);

  // Trend State (N) - Green for bull, Red for bear
  add(`=$N${r0}="BULL"`, C_GREEN, 14);
  add(`=$N${r0}="BEAR"`, C_RED, 14);

  // SMAs (O/P/Q) - Green when price >= SMA (bullish), Red when price < SMA (bearish)
  add(`=AND($O${r0}>0,$G${r0}>=$O${r0})`, C_GREEN, 15);
  add(`=AND($O${r0}>0,$G${r0}<$O${r0})`, C_RED, 15);
  add(`=AND($P${r0}>0,$G${r0}>=$P${r0})`, C_GREEN, 16);
  add(`=AND($P${r0}>0,$G${r0}<$P${r0})`, C_RED, 16);
  add(`=AND($Q${r0}>0,$G${r0}>=$Q${r0})`, C_GREEN, 17);
  add(`=AND($Q${r0}>0,$G${r0}<$Q${r0})`, C_RED, 17);

  // RSI (R) - Green oversold (opportunity), Red overbought
  add(`=$R${r0}<=30`, C_GREEN, 18);
  add(`=$R${r0}>=70`, C_RED, 18);

  // MACD Hist (S) - Green positive, Red negative
  add(`=$S${r0}>0`, C_GREEN, 19);
  add(`=$S${r0}<0`, C_RED, 19);

  // Divergence (T) - Green bullish, Red bearish
  add(`=REGEXMATCH($T${r0},"BULL")`, C_GREEN, 20);
  add(`=REGEXMATCH($T${r0},"BEAR")`, C_RED, 20);

  // ADX (U) - Green strong trend, no red (weak trend stays blue)
  add(`=$U${r0}>=25`, C_GREEN, 21);

  // Stoch %K (V) - Green oversold, Red overbought
  add(`=$V${r0}<=0.2`, C_GREEN, 22);
  add(`=$V${r0}>=0.8`, C_RED, 22);

  // VOL REGIME (W) - Green low vol, Red extreme vol
  add(`=$W${r0}="LOW VOL"`, C_GREEN, 23);
  add(`=$W${r0}="EXTREME VOL"`, C_RED, 23);

  // BBP SIGNAL (X) - Green oversold/mean reversion, Red overbought
  add(`=REGEXMATCH($X${r0},"EXTREME OVERSOLD|MEAN REVERSION")`, C_GREEN, 24);
  add(`=REGEXMATCH($X${r0},"EXTREME OVERBOUGHT")`, C_RED, 24);

  // ATR (Y) - Green low volatility, Red high volatility
  add(`=IFERROR($Y${r0}/$G${r0},0)<=0.02`, C_GREEN, 25);
  add(`=IFERROR($Y${r0}/$G${r0},0)>=0.05`, C_RED, 25);

  // Bollinger %B (Z) - Green oversold, Red overbought
  add(`=$Z${r0}<=0.2`, C_GREEN, 26);
  add(`=$Z${r0}>=0.8`, C_RED, 26);

  // Target (AA) - Green good upside, Red limited upside
  add(`=AND($AA${r0}>0,$AA${r0}>=$G${r0}*1.05)`, C_GREEN, 27);
  add(`=AND($AA${r0}>0,$AA${r0}<=$G${r0}*1.01)`, C_RED, 27);

  // R:R Quality (AB) - Green good R:R, Red poor R:R
  add(`=$AB${r0}>=3`, C_GREEN, 28);
  add(`=$AB${r0}<=1`, C_RED, 28);

  // Support (AC) - Green at/near support, Red below support
  add(`=AND($AC${r0}>0,$G${r0}>=$AC${r0},$G${r0}<=$AC${r0}*1.01)`, C_GREEN, 29);
  add(`=AND($AC${r0}>0,$G${r0}<$AC${r0})`, C_RED, 29);

  // Resistance (AD) - Green far from resistance, Red at resistance
  add(`=AND($AD${r0}>0,$G${r0}<=$AD${r0}*0.90)`, C_GREEN, 30);
  add(`=AND($AD${r0}>0,$G${r0}>=$AD${r0}*0.995)`, C_RED, 30);

  sh.setConditionalFormatRules(rules);
}

/**
 * Apply conditional formatting to market index cells (E2 and H2)
 * Requirements: 10.1, 10.2, 10.3
 * This function should be called AFTER all other formatting operations
 */
function applyMarketIndexConditionalFormatting(sh) {
  if (!sh) return;
  
  // Define colors for positive/negative values
  const C_GREEN = "#C8E6C9";  // Light green (positive)
  const C_RED = "#FFCDD2";    // Light red (negative)
  
  const indexRules = [];
  
  // NIFTY 50 % change (E2) - Green for positive, Red for negative
  indexRules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThan(0)
      .setBackground(C_GREEN)
      .setFontColor("#000000")
      .setRanges([sh.getRange("E2")])
      .build()
  );
  
  indexRules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThan(0)
      .setBackground(C_RED)
      .setFontColor("#000000")
      .setRanges([sh.getRange("E2")])
      .build()
  );
  
  // S&P 500 % change (H2) - Green for positive, Red for negative
  indexRules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThan(0)
      .setBackground(C_GREEN)
      .setFontColor("#000000")
      .setRanges([sh.getRange("H2")])
      .build()
  );
  
  indexRules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThan(0)
      .setBackground(C_RED)
      .setFontColor("#000000")
      .setRanges([sh.getRange("H2")])
      .build()
  );
  
  // Apply index rules by prepending them to existing rules
  // This ensures they take precedence over other formatting
  const existingRules = sh.getConditionalFormatRules();
  sh.setConditionalFormatRules(indexRules.concat(existingRules));
}

/**
 * Apply professional formatting to control rows (Rows 1 and 2)
 * Called AFTER Bloomberg formatting to ensure it's not overwritten
 */
function applyControlRowFormatting(dashboard) {
  if (!dashboard) return;
  
  // Get locale separator
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const SEP = (/^(en|en_)/.test(ss.getSpreadsheetLocale())) ? "," : ";";
  
  // Professional color scheme
  const CONTROL_BG = "#1E3A5F";        // Deep blue
  const CONTROL_LABEL = "#FFD700";     // Gold
  const CONTROL_INPUT = "#2C5282";     // Medium blue
  const SORT_BG = "#0F2942";           // Darker blue
  const NIFTY_BG = "#1A237E";          // Indigo
  const SP500_BG = "#01579B";          // Blue
  
  // Row 1: Country filters (A1-D1)
  dashboard.getRange("A1")
    .setValue("🇺🇸 USA")
    .setBackground(CONTROL_BG).setFontColor(CONTROL_LABEL).setFontWeight("bold")
    .setFontSize(11)
    .setHorizontalAlignment("center").setVerticalAlignment("middle");

  dashboard.getRange("B1")
    .setBackground(CONTROL_INPUT)
    .setHorizontalAlignment("center").setVerticalAlignment("middle");

  dashboard.getRange("C1")
    .setValue("🇮🇳 INDIA")
    .setBackground(CONTROL_BG).setFontColor(CONTROL_LABEL).setFontWeight("bold")
    .setFontSize(11)
    .setHorizontalAlignment("center").setVerticalAlignment("middle");

  dashboard.getRange("D1")
    .setBackground(CONTROL_INPUT)
    .setHorizontalAlignment("center").setVerticalAlignment("middle");

  // Row 1: Category filter (E1-F1)
  dashboard.getRange("E1")
    .setValue("📊 Category")
    .setBackground(CONTROL_BG).setFontColor(CONTROL_LABEL).setFontWeight("bold")
    .setFontSize(11)
    .setHorizontalAlignment("center").setVerticalAlignment("middle");

  dashboard.getRange("F1")
    .setBackground(CONTROL_INPUT).setFontColor("white").setFontWeight("bold")
    .setFontSize(10)
    .setHorizontalAlignment("center").setVerticalAlignment("middle");

  // Row 1: TRADE mode toggle (G1-H1)
  dashboard.getRange("G1")
    .setValue("⚡ INVEST")
    .setBackground(CONTROL_BG).setFontColor(CONTROL_LABEL).setFontWeight("bold")
    .setFontSize(11)
    .setHorizontalAlignment("center").setVerticalAlignment("middle");

  dashboard.getRange("H1")
    .setBackground(CONTROL_INPUT)
    .setHorizontalAlignment("center").setVerticalAlignment("middle");

  // Row 1: Dashboard refresh (I1-J1)
  dashboard.getRange("I1")
    .setValue("🔄 Refresh")
    .setBackground(CONTROL_BG).setFontColor(CONTROL_LABEL).setFontWeight("bold")
    .setFontSize(11)
    .setHorizontalAlignment("center").setVerticalAlignment("middle");

  dashboard.getRange("J1")
    .setBackground(CONTROL_INPUT)
    .setHorizontalAlignment("center").setVerticalAlignment("middle");

  // Row 1: Calculations refresh (K1-L1)
  dashboard.getRange("K1")
    .setValue("🧮 CALC")
    .setBackground(CONTROL_BG).setFontColor(CONTROL_LABEL).setFontWeight("bold")
    .setFontSize(11)
    .setHorizontalAlignment("center").setVerticalAlignment("middle");

  dashboard.getRange("L1")
    .setBackground(CONTROL_INPUT)
    .setHorizontalAlignment("center").setVerticalAlignment("middle");

  // Row 1: Data rebuild (M1-N1)
  dashboard.getRange("M1")
    .setValue("💾 DATA")
    .setBackground(CONTROL_BG).setFontColor(CONTROL_LABEL).setFontWeight("bold")
    .setFontSize(11)
    .setHorizontalAlignment("center").setVerticalAlignment("middle");

  dashboard.getRange("N1")
    .setBackground(CONTROL_INPUT)
    .setHorizontalAlignment("center").setVerticalAlignment("middle");

  // Row 1: Alert mode (O1-P1)
  dashboard.getRange("O1")
    .setValue("🔔 ALERT")
    .setBackground(CONTROL_BG).setFontColor(CONTROL_LABEL).setFontWeight("bold")
    .setFontSize(11)
    .setHorizontalAlignment("center").setVerticalAlignment("middle");

  dashboard.getRange("P1")
    .setBackground(CONTROL_INPUT).setFontColor("white").setFontWeight("bold")
    .setFontSize(10)
    .setHorizontalAlignment("center").setVerticalAlignment("middle");

  // Add borders to Row 1
  dashboard.getRange("A1:P1")
    .setBorder(true, true, true, true, true, true, "#FFD700", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  // Row 2: Sort controls (A2-B2)
  dashboard.getRange("A2")
    .setValue("⬇️ Sort By")
    .setBackground(SORT_BG).setFontColor(CONTROL_LABEL).setFontWeight("bold")
    .setFontSize(10)
    .setHorizontalAlignment("center").setVerticalAlignment("middle");

  dashboard.getRange("B2")
    .setBackground(CONTROL_INPUT).setFontColor("white").setFontWeight("bold")
    .setFontSize(10)
    .setHorizontalAlignment("center").setVerticalAlignment("middle");

  // Row 2: Market indices (C2-H2)
  dashboard.getRange("C2")
    .setValue("NIFTY 50")
    .setBackground(NIFTY_BG).setFontColor("white").setFontWeight("bold")
    .setFontSize(10)
    .setHorizontalAlignment("center").setVerticalAlignment("middle");

  dashboard.getRange("D2")
    .setBackground(NIFTY_BG).setFontColor("white").setFontWeight("bold")
    .setFontSize(10)
    .setHorizontalAlignment("center").setVerticalAlignment("middle")
    .setNumberFormat("#,##0.00");

  dashboard.getRange("E2")
    .setBackground(NIFTY_BG).setFontColor("white").setFontWeight("bold")
    .setFontSize(10)
    .setHorizontalAlignment("center").setVerticalAlignment("middle")
    .setNumberFormat("0.00%");

  dashboard.getRange("F2")
    .setValue("S&P 500")
    .setBackground(SP500_BG).setFontColor("white").setFontWeight("bold")
    .setFontSize(10)
    .setHorizontalAlignment("center").setVerticalAlignment("middle");

  dashboard.getRange("G2")
    .setBackground(SP500_BG).setFontColor("white").setFontWeight("bold")
    .setFontSize(10)
    .setHorizontalAlignment("center").setVerticalAlignment("middle")
    .setNumberFormat("#,##0.00");

  dashboard.getRange("H2")
    .setBackground(SP500_BG).setFontColor("white").setFontWeight("bold")
    .setFontSize(10)
    .setHorizontalAlignment("center").setVerticalAlignment("middle")
    .setNumberFormat("0.00%");

  // Add borders to Row 2
  dashboard.getRange("A2:H2")
    .setBorder(true, true, true, true, true, true, "#4A90E2", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  // Set row heights for better appearance
  dashboard.setRowHeight(1, 28);
  dashboard.setRowHeight(2, 26);
}

function applyDashboardGroupMapAndColors_(sh) {
  if (!sh) return;

  const COLORS = {
    IDENTITY: "#37474F",
    SIGNALING: "#1565C0",
    PRICE_VOLUME: "#D84315",    // Deep Orange
    PERFORMANCE: "#1976D2",     // Medium Blue
    TREND: "#00838F",           // Cyan
    MOMENTUM: "#F57C00",        // Orange
    VOLATILITY: "#C62828",      // Red
    TARGET: "#AD1457"           // Pink (includes all target-related columns Y-AE)
  };

  const FG = "#FFFFFF";

  // GROUPS array for 33 columns (A-AG) - Matches new structure
  const GROUPS = [
    { name: "IDENTITY", c1: 1, c2: 1, color: COLORS.IDENTITY },           // A
    { name: "SIGNALING", c1: 2, c2: 6, color: COLORS.SIGNALING },         // B-F (MARKET RATING, DECISION, SIGNAL, PATTERNS, CONSENSUS PRICE)
    { name: "PRICE / VOLUME", c1: 7, c2: 9, color: COLORS.PRICE_VOLUME }, // G-I (Price, Change%, Vol Trend)
    { name: "PERFORMANCE", c1: 10, c2: 13, color: COLORS.PERFORMANCE },   // J-M (ATH TRUE, ATH Diff%, ATH ZONE, FUNDAMENTAL)
    { name: "TREND", c1: 14, c2: 17, color: COLORS.TREND },               // N-Q (Trend State, SMA 20/50/200)
    { name: "MOMENTUM", c1: 18, c2: 22, color: COLORS.MOMENTUM },         // R-V (RSI, MACD, Div, ADX, Stoch)
    { name: "VOLATILITY", c1: 23, c2: 26, color: COLORS.VOLATILITY },     // W-Z (VOL REGIME, BBP SIGNAL, ATR, Bollinger %B)
    { name: "TARGET", c1: 27, c2: 33, color: COLORS.TARGET }              // AA-AG (Target, R:R, Support, Res, ATR STOP/TARGET, Position)
  ];

  const style = (row, c1, c2, bg) => {
    sh.getRange(row, c1, 1, c2 - c1 + 1)
      .setBackground(bg)
      .setFontColor(FG)
      .setFontWeight("bold")
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle")
      .setWrap(true);
  };

  // Clear all existing merges in row 3 first to avoid conflicts
  try {
    sh.getRange(3, 1, 1, 33).breakApart();
  } catch (e) {
    // Ignore if no merges exist
  }

  GROUPS.forEach(g => {
    // Style row 3 (group headers)
    style(3, g.c1, g.c2, g.color);
    const r3 = sh.getRange(3, g.c1, 1, g.c2 - g.c1 + 1);
    if (g.c1 !== g.c2) r3.merge();
    r3.setValue(g.name);
    
    // Style row 4 (column headers) with same group color
    style(4, g.c1, g.c2, g.color);
  });
}


// ============================================================================
// CONTROL ROW FUNCTIONS
// ============================================================================

/**
 * Update country filter based on B1 and D1 checkboxes
 */
function updateCountryFilter() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dashboard = ss.getSheetByName('DASHBOARD');
  
  if (!dashboard) return;
  
  const usaChecked = dashboard.getRange('B1').getValue();
  const indiaChecked = dashboard.getRange('D1').getValue();
  
  let countryFilter = [];
  if (usaChecked) countryFilter.push('USA');
  if (indiaChecked) countryFilter.push('INDIA');  // Changed from 'IN' to 'INDIA'
  
  // If both unchecked, show all
  if (countryFilter.length === 0) {
    countryFilter = ['ALL'];
  }
  
  // Removed: Write to INPUT!B1 - data persists in DASHBOARD only
  
  // Refresh dashboard data to apply filter (preserve checkbox states)
  SpreadsheetApp.flush();
  const DATA_START_ROW = 5;
  refreshDashboardData(dashboard, ss, DATA_START_ROW, true);
  
  ss.toast('Country filter updated', '✓ Filter', 2);
}

/**
 * Update category filter based on F1 dropdown
 */
function updateCategoryFilter() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dashboard = ss.getSheetByName('DASHBOARD');
  
  if (!dashboard) return;
  
  const categoryValue = dashboard.getRange('F1').getValue();
  
  // Removed: Write to INPUT!C1 - data persists in DASHBOARD only
  
  // Refresh dashboard data to apply filter (preserve checkbox states)
  SpreadsheetApp.flush();
  const DATA_START_ROW = 5;
  refreshDashboardData(dashboard, ss, DATA_START_ROW, true);
  
  ss.toast('Category filter updated', '✓ Filter', 2);
}

/**
 * Sync mode toggle between DASHBOARD H1 and INPUT G1
 */
function syncModeToggle(sourceSheet, sourceCell) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dashboard = ss.getSheetByName('DASHBOARD');
  
  if (!dashboard) return;
  
  // Prevent infinite loop
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('SYNCING_MODE') === 'true') {
    return;
  }
  
  props.setProperty('SYNCING_MODE', 'true');
  
  try {
    const modeValue = dashboard.getRange('H1').getValue();
    
    // Removed: Sync to INPUT!G1 - data persists in DASHBOARD only
    
    // Update signal formulas
    updateSignalFormulas();
    
    // Show notification
    const mode = modeValue ? 'INVEST' : 'TRADE';
    ss.toast(`Switched to ${mode} mode`, '✓ Mode Updated', 3);
    
  } finally {
    props.deleteProperty('SYNCING_MODE');
  }
}

/**
 * Refresh dashboard data without rebuilding layout (wrapper for J1 checkbox)
 */
function refreshDashboardDataFromCheckbox() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dashboard = ss.getSheetByName('DASHBOARD');
  
  if (!dashboard) return;
  
  ss.toast('Refreshing dashboard data...', '⚙️ Processing', -1);
  
  try {
    const DATA_START_ROW = 5;
    // Preserve checkbox states when manually refreshing
    refreshDashboardData(dashboard, ss, DATA_START_ROW, true);
    
    // Reapply current sort
    const sortColumn = dashboard.getRange('B2').getValue();
    if (sortColumn && sortColumn !== 'Change %') {
      sortDashboardByColumn(sortColumn);
    }
    
    ss.toast('Dashboard refreshed successfully', '✓ Complete', 3);
    
  } catch (error) {
    ss.toast(`Error: ${error.message}`, '⚠️ Refresh Failed', 5);
  } finally {
    // Reset only the refresh checkbox
    dashboard.getRange('J1').setValue(false);
  }
}

/**
 * Refresh calculations sheet
 */
function refreshCalculations() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dashboard = ss.getSheetByName('DASHBOARD');
  const calculations = ss.getSheetByName('CALCULATIONS');
  
  if (!dashboard || !calculations) return;
  
  ss.toast('Refreshing calculations...', '⚙️ Processing', -1);
  
  try {
    // Force recalculation
    calculations.getDataRange().activate();
    SpreadsheetApp.flush();
    
    ss.toast('Calculations refreshed successfully', '✓ Complete', 3);
    
  } catch (error) {
    ss.toast(`Error: ${error.message}`, '⚠️ Refresh Failed', 5);
  } finally {
    // Reset checkbox
    dashboard.getRange('L1').setValue(false);
  }
}

/**
 * Rebuild DATA sheet
 */
function rebuildDataSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dashboard = ss.getSheetByName('DASHBOARD');
  
  if (!dashboard) return;
  
  ss.toast('Rebuilding DATA sheet...', '⚙️ Processing', -1);
  
  try {
    // Call existing function
    FlushDataSheetAndBuild();
    
    ss.toast('DATA sheet rebuilt successfully', '✓ Complete', 3);
    
  } catch (error) {
    ss.toast(`Error: ${error.message}`, '⚠️ Rebuild Failed', 5);
  } finally {
    // Reset checkbox
    dashboard.getRange('N1').setValue(false);
  }
}

/**
 * Sort dashboard by selected column
 */
function sortDashboardByColumn(columnName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dashboard = ss.getSheetByName('DASHBOARD');
  
  if (!dashboard) return;
  
  try {
    // Find column index from Row 4 headers
    const headers = dashboard.getRange('A4:AG4').getValues()[0];
    const columnIndex = headers.indexOf(columnName);
    
    if (columnIndex === -1) {
      throw new Error(`Column "${columnName}" not found`);
    }
    
    // Column index is 0-based, convert to 1-based
    const sortColumnNumber = columnIndex + 1;
    
    // Get data range (Row 5 onwards)
    const lastRow = dashboard.getLastRow();
    
    if (lastRow < 5) {
      return; // No data to sort
    }
    
    const dataRange = dashboard.getRange(5, 1, lastRow - 4, 33);
    
    // Sort descending (highest to lowest)
    dataRange.sort({
      column: sortColumnNumber,
      ascending: false
    });
    
    SpreadsheetApp.flush();
    
  } catch (error) {
    ss.toast(`Sort error: ${error.message}`, '⚠️ Sort Failed', 5);
  }
}

/**
 * Handle sort column change - refresh filter formula with new sort column
 */
function onSortColumnChange() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dashboard = ss.getSheetByName('DASHBOARD');
  
  if (!dashboard) return;
  
  const columnName = dashboard.getRange('B2').getValue();
  
  if (columnName) {
    // Refresh dashboard data with new sort column (preserve checkbox states)
    const DATA_START_ROW = 5;
    refreshDashboardData(dashboard, ss, DATA_START_ROW, true);
    ss.toast(`Sorted by ${columnName}`, '✓ Sort', 2);
  }
}

/**
 * Setup dropdown data validations for control row
 */
function setupControlRowDropdowns() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dashboard = ss.getSheetByName('DASHBOARD');
  const input = ss.getSheetByName('INPUT');
  
  if (!dashboard || !input) return;
  
  // Setup Category multi-select dropdown (F1) - read from H2:J15 (3 columns)
  const categoryData = input.getRange('H2:J15').getValues();
  const categories = [];
  
  // Flatten and collect all non-empty values from the 3 columns
  categoryData.forEach(row => {
    row.forEach(cell => {
      if (cell !== '') {
        categories.push(cell);
      }
    });
  });
  
  // Add "ALL" option at the beginning
  categories.unshift('ALL');
  
  // For multi-select, we need to manually enable it through data validation
  // Create a simple list dropdown first
  const categoryRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(categories, true)
    .setAllowInvalid(true)
    .build();
  
  dashboard.getRange('F1').setDataValidation(categoryRule);
  
  // Add note with instructions for multi-select
  dashboard.getRange('F1').setNote(
    '💡 For multi-select:\n' +
    '1. Right-click on this cell\n' +
    '2. Select "Data validation"\n' +
    '3. Check "Show dropdown list in cell"\n' +
    '4. Check "Show checkboxes"\n' +
    '\n' +
    'Or manually enter comma-separated values:\n' +
    'Example: Tech, Finance, Healthcare'
  );
  
  // Setup Alert dropdown (P1)
  const alertLevels = ['All', 'HIGH', 'CRITICAL'];
  
  const alertRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(alertLevels, true)
    .setAllowInvalid(false)
    .build();
  
  dashboard.getRange('P1').setDataValidation(alertRule);
  
  // Setup Sort dropdown (B2)
  const headers = dashboard.getRange('B4:AG4').getValues()[0]
    .filter(h => h !== '');
  
  const sortRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(headers, true)
    .setAllowInvalid(false)
    .build();
  
  dashboard.getRange('B2').setDataValidation(sortRule);
  
  ss.toast('Control row dropdowns configured\n\nNote: For Category multi-select, right-click F1 > Data validation > Enable checkboxes', '✓ Setup', 5);
}

/**
 * Get alert severity filter for Monitor.js integration
 */
function getAlertSeverityFilter() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dashboard = ss.getSheetByName('DASHBOARD');
  
  if (!dashboard) return 'All';
  
  return dashboard.getRange('P1').getValue() || 'All';
}
