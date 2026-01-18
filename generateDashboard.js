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
    
    const DATA_START_ROW = 4;
    const SENTINEL = "DASHBOARD_LAYOUT_V1_BLOOMBERG";
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

  // Row 1 controls
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

  // Market indices in row 1
  dashboard.getRange("E1")
    .setValue("NIFTY 50")
    .setBackground("#1A237E").setFontColor("white").setFontWeight("bold")
    .setHorizontalAlignment("center").setVerticalAlignment("middle");

  dashboard.getRange("F1")
    .setFormula(`=IFERROR(GOOGLEFINANCE("INDEXNSE:NIFTY_50"${SEP}"price")${SEP}0)`)
    .setBackground("#1A237E").setFontColor("white").setFontWeight("bold")
    .setHorizontalAlignment("center").setVerticalAlignment("middle")
    .setNumberFormat("#,##0.00");

  dashboard.getRange("G1")
    .setFormula(`=IFERROR(GOOGLEFINANCE("INDEXNSE:NIFTY_50"${SEP}"changepct")/100${SEP}0)`)
    .setBackground("#1A237E").setFontColor("white").setFontWeight("bold")
    .setHorizontalAlignment("center").setVerticalAlignment("middle")
    .setNumberFormat("0.00%");

  dashboard.getRange("H1")
    .setValue("S&P 500")
    .setBackground("#01579B").setFontColor("white").setFontWeight("bold")
    .setHorizontalAlignment("center").setVerticalAlignment("middle");

  dashboard.getRange("I1")
    .setFormula(`=IFERROR(GOOGLEFINANCE("INDEXSP:.INX"${SEP}"price")${SEP}0)`)
    .setBackground("#01579B").setFontColor("white").setFontWeight("bold")
    .setHorizontalAlignment("center").setVerticalAlignment("middle")
    .setNumberFormat("#,##0.00");

  dashboard.getRange("J1")
    .setFormula(`=IFERROR(GOOGLEFINANCE("INDEXSP:.INX"${SEP}"changepct")/100${SEP}0)`)
    .setBackground("#01579B").setFontColor("white").setFontWeight("bold")
    .setHorizontalAlignment("center").setVerticalAlignment("middle")
    .setNumberFormat("0.00%");

  // Row 2 group headers (updated structure to match CALCULATIONS)
  const styleGroup = (a1, label, bg) => {
    dashboard.getRange(a1).merge()
      .setValue(label)
      .setBackground(bg).setFontColor("white").setFontWeight("bold")
      .setHorizontalAlignment("center").setVerticalAlignment("middle");
  };

  // Clear any existing merges in row 2 to avoid merge conflicts
  try {
    dashboard.getRange("A2:AE2").breakApart();
  } catch (e) {
    // Ignore if no merges exist
  }

  dashboard.getRange("A2:AE2").clearContent();
  styleGroup("A2:A2", "IDENTITY", "#37474F");        // Dark Blue-Grey (A)
  styleGroup("B2:D2", "SIGNALING", "#1565C0");       // Blue (B-D: SIGNAL, PATTERNS, DECISION)
  styleGroup("E2:G2", "PRICE / VOLUME", "#D84315");  // Deep Orange (E-G: Price, Change%, Vol Trend)
  styleGroup("H2:K2", "PERFORMANCE", "#1976D2");     // Medium Blue (H-K: ATH TRUE, ATH Diff%, ATH ZONE, FUNDAMENTAL)
  styleGroup("L2:O2", "TREND", "#00838F");           // Cyan (L-O: Trend State, SMA 20/50/200)
  styleGroup("P2:T2", "MOMENTUM", "#F57C00");        // Orange (P-T: RSI, MACD, Div, ADX, Stoch)
  styleGroup("U2:X2", "VOLATILITY", "#C62828");      // Red (U-X: VOL REGIME, BBP SIGNAL, ATR, Bollinger %B)
  styleGroup("Y2:AE2", "TARGET", "#AD1457");         // Pink (Y-AE: All target-related) - MATCHES CALCULATIONS
  dashboard.getRange("A2:AE2").setWrap(true);

  // Row 3 column headers - MATCHES CALCULATIONS ORDER EXACTLY
  const headers = [[
    "Ticker",           // A
    "SIGNAL",           // B - MATCHES CALCULATIONS
    "PATTERNS",         // C - MATCHES CALCULATIONS
    "DECISION",         // D - MATCHES CALCULATIONS
    "Price",            // E
    "Change %",         // F
    "Vol Trend",        // G
    "ATH (TRUE)",       // H
    "ATH Diff %",       // I
    "ATH ZONE",         // J
    "FUNDAMENTAL",      // K - MATCHES CALCULATIONS
    "Trend State",      // L
    "SMA 20",           // M
    "SMA 50",           // N
    "SMA 200",          // O
    "RSI",              // P
    "MACD Hist",        // Q
    "Divergence",       // R
    "ADX (14)",         // S
    "Stoch %K (14)",    // T
    "VOL REGIME",       // U - MATCHES CALCULATIONS
    "BBP SIGNAL",       // V - MATCHES CALCULATIONS
    "ATR (14)",         // W - MATCHES CALCULATIONS
    "Bollinger %B",     // X - MATCHES CALCULATIONS
    "Target (3:1)",     // Y
    "R:R Quality",      // Z
    "Support",          // AA
    "Resistance",       // AB
    "ATR STOP",         // AC
    "ATR TARGET",       // AD
    "POSITION SIZE"     // AE (31 columns total)
  ]];

  dashboard.getRange(3, 1, 1, 31)
    .setValues(headers)
    .setBackground("#0D0D0D").setFontColor("#FFD700").setFontWeight("bold")
    .setHorizontalAlignment("center").setVerticalAlignment("middle")
    .setWrap(true);

  // Freeze panes
  dashboard.setFrozenRows(3);
  dashboard.setFrozenColumns(1);

  // White border for top header rows
  dashboard.getRange("A1:AE3")
    .setBorder(true, true, true, true, true, true, "#FFFFFF", SpreadsheetApp.BorderStyle.SOLID);

  // Sentinel note
  dashboard.getRange("A1").setNote(SENTINEL);
}

function refreshDashboardData(dashboard, ss, DATA_START_ROW) {
  // Clear existing data (31 columns A-AE, includes R:R Quality)
  dashboard.getRange(DATA_START_ROW, 1, 1000, 31).clearContent();

  // Reset checkboxes to false
  dashboard.getRange("B1").setValue(false);
  dashboard.getRange("D1").setValue(false);

  // Filter formula - pulls columns from CALCULATIONS in CORRECT ORDER (matches CALCULATIONS exactly)
  const filterFormula =
    '=IFERROR(' +
    'SORT(' +
    'FILTER({' +
    'CALCULATIONS!$A$3:$A,' +  // A: Ticker
    'CALCULATIONS!$B$3:$B,' +  // B: SIGNAL - MATCHES CALCULATIONS
    'CALCULATIONS!$C$3:$C,' +  // C: PATTERNS - MATCHES CALCULATIONS
    'CALCULATIONS!$D$3:$D,' +  // D: DECISION - MATCHES CALCULATIONS
    'CALCULATIONS!$E$3:$E,' +  // E: Price
    'CALCULATIONS!$F$3:$F,' +  // F: Change %
    'CALCULATIONS!$G$3:$G,' +  // G: Vol Trend
    'CALCULATIONS!$H$3:$H,' +  // H: ATH (TRUE)
    'CALCULATIONS!$I$3:$I,' +  // I: ATH Diff %
    'CALCULATIONS!$J$3:$J,' +  // J: ATH ZONE
    'CALCULATIONS!$K$3:$K,' +  // K: FUNDAMENTAL
    'CALCULATIONS!$L$3:$L,' +  // L: Trend State
    'CALCULATIONS!$M$3:$M,' +  // M: SMA 20
    'CALCULATIONS!$N$3:$N,' +  // N: SMA 50
    'CALCULATIONS!$O$3:$O,' +  // O: SMA 200
    'CALCULATIONS!$P$3:$P,' +  // P: RSI
    'CALCULATIONS!$Q$3:$Q,' +  // Q: MACD Hist
    'CALCULATIONS!$R$3:$R,' +  // R: Divergence
    'CALCULATIONS!$S$3:$S,' +  // S: ADX (14)
    'CALCULATIONS!$T$3:$T,' +  // T: Stoch %K (14)
    'CALCULATIONS!$U$3:$U,' +  // U: VOL REGIME - MATCHES CALCULATIONS
    'CALCULATIONS!$V$3:$V,' +  // V: BBP SIGNAL - MATCHES CALCULATIONS
    'CALCULATIONS!$W$3:$W,' +  // W: ATR (14) - MATCHES CALCULATIONS
    'CALCULATIONS!$X$3:$X,' +  // X: Bollinger %B - MATCHES CALCULATIONS
    'CALCULATIONS!$Y$3:$Y,' +  // Y: Target (3:1)
    'CALCULATIONS!$Z$3:$Z,' +  // Z: R:R Quality
    'CALCULATIONS!$AA$3:$AA,' + // AA: Support
    'CALCULATIONS!$AB$3:$AB,' + // AB: Resistance
    'CALCULATIONS!$AC$3:$AC,' + // AC: ATR STOP
    'CALCULATIONS!$AD$3:$AD,' + // AD: ATR TARGET
    'CALCULATIONS!$AE$3:$AE' +  // AE: POSITION SIZE
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

  // Apply Bloomberg formatting + heatmap
  applyDashboardBloombergFormatting_(dashboard, DATA_START_ROW);
  
  // Apply group colors AFTER Bloomberg formatting to ensure they're not overwritten
  applyDashboardGroupMapAndColors_(dashboard);
}

function applyDashboardBloombergFormatting_(sh, DATA_START_ROW) {
  if (!sh) return;

  const C_BLUE = "#E3F2FD";   // Light blue (default background)
  const C_GREEN = "#C8E6C9";  // Light green (positive)
  const C_RED = "#FFCDD2";    // Light red (negative)
  const HEADER_DARK = "#1F1F1F";
  const TOTAL_COLS = 31; // Updated to 31 columns (A-AE, includes R:R Quality)

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
  const dataRange = sh.getRange(DATA_START_ROW, 1, numRows, 31); // Updated to 31 columns (includes R:R Quality)

  // Header styling
  sh.getRange(1, 1, 1, TOTAL_COLS)
    .setBackground(HEADER_DARK)
    .setFontColor("#FFFFFF")
    .setFontWeight("bold")
    .setVerticalAlignment("middle");

  sh.getRange("A1:D1").setHorizontalAlignment("center");
  sh.getRange("E1:J1").setHorizontalAlignment("center");

  sh.getRange(2, 1, 1, TOTAL_COLS)
    .setBackground("#E7E6E6")
    .setFontColor("#000000")
    .setFontWeight("bold")
    .setVerticalAlignment("middle")
    .setHorizontalAlignment("center");

  sh.getRange(3, 1, 1, TOTAL_COLS)
    .setBackground(HEADER_DARK)
    .setFontColor("#FFFFFF")
    .setFontWeight("bold")
    .setVerticalAlignment("middle")
    .setHorizontalAlignment("center")
    .setWrap(true);

  // Column widths (31 columns)
  for (let c = 1; c <= 31; c++) sh.setColumnWidth(c, 85);

  // Row heights
  sh.setRowHeight(1, 22);
  sh.setRowHeight(2, 18);
  sh.setRowHeight(3, 22);
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

  sh.getRange(1, 1, 3, TOTAL_COLS)
    .setBorder(true, true, true, true, true, true, "#FFFFFF", SpreadsheetApp.BorderStyle.SOLID);

  // Number formats (updated for correct column structure)
  sh.getRange(DATA_START_ROW, 5, numRows, 1).setNumberFormat("#,##0.00");  // E: Price
  sh.getRange(DATA_START_ROW, 6, numRows, 1).setNumberFormat("0.00%");     // F: Change%
  sh.getRange(DATA_START_ROW, 7, numRows, 1).setNumberFormat("0.00");      // G: Vol Trend (RVOL)
  sh.getRange(DATA_START_ROW, 8, numRows, 1).setNumberFormat("#,##0.00");  // H: ATH (TRUE)
  sh.getRange(DATA_START_ROW, 9, numRows, 1).setNumberFormat("0.00%");     // I: ATH Diff%
  sh.getRange(DATA_START_ROW, 10, numRows, 1).setNumberFormat("@");        // J: ATH ZONE
  sh.getRange(DATA_START_ROW, 11, numRows, 1).setNumberFormat("@");        // K: FUNDAMENTAL
  sh.getRange(DATA_START_ROW, 13, numRows, 3).setNumberFormat("#,##0.00"); // M-O: SMAs
  sh.getRange(DATA_START_ROW, 16, numRows, 1).setNumberFormat("0.0");      // P: RSI
  sh.getRange(DATA_START_ROW, 17, numRows, 1).setNumberFormat("0.000");    // Q: MACD
  sh.getRange(DATA_START_ROW, 19, numRows, 1).setNumberFormat("0.0");      // S: ADX
  sh.getRange(DATA_START_ROW, 20, numRows, 1).setNumberFormat("0.00%");    // T: Stoch
  sh.getRange(DATA_START_ROW, 21, numRows, 1).setNumberFormat("@");        // U: VOL REGIME
  sh.getRange(DATA_START_ROW, 22, numRows, 1).setNumberFormat("@");        // V: BBP SIGNAL
  sh.getRange(DATA_START_ROW, 23, numRows, 1).setNumberFormat("#,##0.00"); // W: ATR
  sh.getRange(DATA_START_ROW, 24, numRows, 1).setNumberFormat("0.00");     // X: Bollinger %B
  sh.getRange(DATA_START_ROW, 25, numRows, 1).setNumberFormat("#,##0.00"); // Y: Target (3:1)
  sh.getRange(DATA_START_ROW, 26, numRows, 1).setNumberFormat("0.00");     // Z: R:R Quality
  sh.getRange(DATA_START_ROW, 27, numRows, 1).setNumberFormat("#,##0.00"); // AA: Support
  sh.getRange(DATA_START_ROW, 28, numRows, 1).setNumberFormat("#,##0.00"); // AB: Resistance
  sh.getRange(DATA_START_ROW, 29, numRows, 1).setNumberFormat("#,##0.00"); // AC: ATR STOP
  sh.getRange(DATA_START_ROW, 30, numRows, 1).setNumberFormat("#,##0.00"); // AD: ATR TARGET
  sh.getRange(DATA_START_ROW, 31, numRows, 1).setNumberFormat("@");        // AE: POSITION SIZE

  // Conditional formatting for market indices in row 1
  const indexRules = [];
  
  // NIFTY 50 % change (G1) - Green for positive, Red for negative
  indexRules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThan(0)
      .setBackground(C_GREEN)
      .setFontColor("#000000")
      .setRanges([sh.getRange("G1")])
      .build()
  );
  
  indexRules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThan(0)
      .setBackground(C_RED)
      .setFontColor("#000000")
      .setRanges([sh.getRange("G1")])
      .build()
  );
  
  // S&P 500 % change (J1) - Green for positive, Red for negative
  indexRules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThan(0)
      .setBackground(C_GREEN)
      .setFontColor("#000000")
      .setRanges([sh.getRange("J1")])
      .build()
  );
  
  indexRules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThan(0)
      .setBackground(C_RED)
      .setFontColor("#000000")
      .setRanges([sh.getRange("J1")])
      .build()
  );
  
  // Apply index rules
  sh.setConditionalFormatRules(indexRules.concat(sh.getConditionalFormatRules()));

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

  // SIGNAL (B) - Green for bullish, Red for bearish
  add(`=REGEXMATCH($B${r0},"STRONG BUY|ATH BREAKOUT|VOLATILITY BREAKOUT|BUY|ACCUMULATE|BREAKOUT|MOMENTUM|UPTREND|BULLISH|OVERSOLD")`, C_GREEN, 2);
  add(`=REGEXMATCH($B${r0},"STOP OUT|RISK OFF")`, C_RED, 2);

  // PATTERNS (C) - Green for bullish patterns, Red for bearish
  add(`=REGEXMATCH($C${r0},"ASC_TRI|BRKOUT|DBL_BTM|INV_H&S|CUP_HDL")`, C_GREEN, 3);
  add(`=REGEXMATCH($C${r0},"DESC_TRI|H&S|DBL_TOP")`, C_RED, 3);

  // DECISION (D) - Green for buy signals, Red for sell signals
  add(`=REGEXMATCH($D${r0},"STRONG BUY|BUY|ADD|STRONG TRADE|TRADE LONG|Accumulate|Add in Dip")`, C_GREEN, 4);
  add(`=REGEXMATCH($D${r0},"EXIT|AVOID|STOP OUT|Stop-Out|Risk-Off|Take Profit|TRIM")`, C_RED, 4);

  // PRICE (E) and Change% (F) - Green for positive, Red for negative
  add(`=$F${r0}>0`, C_GREEN, 5);
  add(`=$F${r0}<0`, C_RED, 5);
  add(`=$F${r0}>0`, C_GREEN, 6);
  add(`=$F${r0}<0`, C_RED, 6);

  // Vol Trend RVOL (G) - Green for high volume, Red for low
  add(`=$G${r0}>=1.5`, C_GREEN, 7);
  add(`=$G${r0}<=0.85`, C_RED, 7);

  // ATH (TRUE) (H) - Green near ATH, Red far from ATH
  add(`=AND($H${r0}>0,$E${r0}>=$H${r0}*0.995)`, C_GREEN, 8);
  add(`=AND($H${r0}>0,$E${r0}<=$H${r0}*0.80)`, C_RED, 8);

  // ATH Diff % (I) - Green near ATH, Red far from ATH
  add(`=$I${r0}>=-0.05`, C_GREEN, 9);
  add(`=$I${r0}<=-0.20`, C_RED, 9);

  // ATH ZONE (J) - Green at/near ATH, Red in correction
  add(`=REGEXMATCH($J${r0},"AT ATH|NEAR ATH")`, C_GREEN, 10);
  add(`=REGEXMATCH($J${r0},"DEEP VALUE|CORRECTION")`, C_RED, 10);

  // FUNDAMENTAL (K) - Green for value, Red for expensive
  add(`=$K${r0}="VALUE"`, C_GREEN, 11);
  add(`=REGEXMATCH($K${r0},"EXPENSIVE|PRICED FOR PERFECTION|ZOMBIE")`, C_RED, 11);

  // Trend State (L) - Green for bull, Red for bear
  add(`=$L${r0}="BULL"`, C_GREEN, 12);
  add(`=$L${r0}="BEAR"`, C_RED, 12);

  // SMAs (M/N/O) - Green above SMA, Red below
  add(`=AND($M${r0}>0,$E${r0}>=$M${r0})`, C_GREEN, 13);
  add(`=AND($M${r0}>0,$E${r0}<$M${r0})`, C_RED, 13);
  add(`=AND($N${r0}>0,$E${r0}>=$N${r0})`, C_GREEN, 14);
  add(`=AND($N${r0}>0,$E${r0}<$N${r0})`, C_RED, 14);
  add(`=AND($O${r0}>0,$E${r0}>=$O${r0})`, C_GREEN, 15);
  add(`=AND($O${r0}>0,$E${r0}<$O${r0})`, C_RED, 15);

  // RSI (P) - Green oversold (opportunity), Red overbought
  add(`=$P${r0}<=30`, C_GREEN, 16);
  add(`=$P${r0}>=70`, C_RED, 16);

  // MACD Hist (Q) - Green positive, Red negative
  add(`=$Q${r0}>0`, C_GREEN, 17);
  add(`=$Q${r0}<0`, C_RED, 17);

  // Divergence (R) - Green bullish, Red bearish
  add(`=REGEXMATCH($R${r0},"BULL")`, C_GREEN, 18);
  add(`=REGEXMATCH($R${r0},"BEAR")`, C_RED, 18);

  // ADX (S) - Green strong trend, no red (weak trend stays blue)
  add(`=$S${r0}>=25`, C_GREEN, 19);

  // Stoch %K (T) - Green oversold, Red overbought
  add(`=$T${r0}<=0.2`, C_GREEN, 20);
  add(`=$T${r0}>=0.8`, C_RED, 20);

  // VOL REGIME (U) - Green low vol, Red extreme vol
  add(`=$U${r0}="LOW VOL"`, C_GREEN, 21);
  add(`=$U${r0}="EXTREME VOL"`, C_RED, 21);

  // BBP SIGNAL (V) - Green oversold/mean reversion, Red overbought
  add(`=REGEXMATCH($V${r0},"EXTREME OVERSOLD|MEAN REVERSION")`, C_GREEN, 22);
  add(`=REGEXMATCH($V${r0},"EXTREME OVERBOUGHT")`, C_RED, 22);

  // ATR (W) - Green low volatility, Red high volatility
  add(`=IFERROR($W${r0}/$E${r0},0)<=0.02`, C_GREEN, 23);
  add(`=IFERROR($W${r0}/$E${r0},0)>=0.05`, C_RED, 23);

  // Bollinger %B (X) - Green oversold, Red overbought
  add(`=$X${r0}<=0.2`, C_GREEN, 24);
  add(`=$X${r0}>=0.8`, C_RED, 24);

  // Target (Y) - Green good upside, Red limited upside
  add(`=AND($Y${r0}>0,$Y${r0}>=$E${r0}*1.05)`, C_GREEN, 25);
  add(`=AND($Y${r0}>0,$Y${r0}<=$E${r0}*1.01)`, C_RED, 25);

  // R:R Quality (Z) - Green good R:R, Red poor R:R
  add(`=$Z${r0}>=3`, C_GREEN, 26);
  add(`=$Z${r0}<=1`, C_RED, 26);

  // Support (AA) - Green at/near support, Red below support
  add(`=AND($AA${r0}>0,$E${r0}>=$AA${r0},$E${r0}<=$AA${r0}*1.01)`, C_GREEN, 27);
  add(`=AND($AA${r0}>0,$E${r0}<$AA${r0})`, C_RED, 27);

  // Resistance (AB) - Green far from resistance, Red at resistance
  add(`=AND($AB${r0}>0,$E${r0}<=$AB${r0}*0.90)`, C_GREEN, 28);
  add(`=AND($AB${r0}>0,$E${r0}>=$AB${r0}*0.995)`, C_RED, 28);

  sh.setConditionalFormatRules(rules);
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

  // GROUPS array for 31 columns (A-AE) - Matches CALCULATIONS structure
  const GROUPS = [
    { name: "IDENTITY", c1: 1, c2: 1, color: COLORS.IDENTITY },           // A
    { name: "SIGNALING", c1: 2, c2: 4, color: COLORS.SIGNALING },         // B-D (SIGNAL, PATTERNS, DECISION)
    { name: "PRICE / VOLUME", c1: 5, c2: 7, color: COLORS.PRICE_VOLUME }, // E-G (Price, Change%, Vol Trend)
    { name: "PERFORMANCE", c1: 8, c2: 11, color: COLORS.PERFORMANCE },    // H-K (ATH TRUE, ATH Diff%, ATH ZONE, FUNDAMENTAL)
    { name: "TREND", c1: 12, c2: 15, color: COLORS.TREND },               // L-O (Trend State, SMA 20/50/200)
    { name: "MOMENTUM", c1: 16, c2: 20, color: COLORS.MOMENTUM },         // P-T (RSI, MACD, Div, ADX, Stoch)
    { name: "VOLATILITY", c1: 21, c2: 24, color: COLORS.VOLATILITY },     // U-X (VOL REGIME, BBP SIGNAL, ATR, Bollinger %B)
    { name: "TARGET", c1: 25, c2: 31, color: COLORS.TARGET }              // Y-AE (All target-related: Target, R:R, Support, Res, ATR STOP/TARGET, Position)
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

  // Clear all existing merges in row 2 first to avoid conflicts
  try {
    sh.getRange(2, 1, 1, 31).breakApart();
  } catch (e) {
    // Ignore if no merges exist
  }

  GROUPS.forEach(g => {
    // Style row 2 (group headers)
    style(2, g.c1, g.c2, g.color);
    const r2 = sh.getRange(2, g.c1, 1, g.c2 - g.c1 + 1);
    if (g.c1 !== g.c2) r2.merge();
    r2.setValue(g.name);
    
    // Style row 3 (column headers) with same group color
    style(3, g.c1, g.c2, g.color);
  });
}

// ============================================================================
// CALCULATIONS SHEET COLUMNS (33 columns: A-AG)
// ============================================================================

const CALC_COLUMNS = {
  // IDENTITY
  TICKER: 'A',
  
  // SIGNALING (B-D)
  SIGNAL: 'B',
  PATTERNS: 'C',        // NEW: Moved from AF
  DECISION: 'D',
  
  // PRICE / VOLUME (E-G) - Only actual volume indicator
  PRICE: 'E',
  CHANGE_PCT: 'F',
  VOL_TREND: 'G',       // RVOL - the only true volume indicator
  
  // PERFORMANCE (H-J)
  ATH_TRUE: 'H',
  ATH_DIFF_PCT: 'I',
  ATH_ZONE: 'J',        // NEW: Moved from K
  
  // FUNDAMENTAL (K)
  FUNDAMENTAL: 'K',     // NEW: Moved from L
  
  // TREND (L-O) - TREND_SCORE removed
  TREND_STATE: 'L',
  SMA_20: 'M',
  SMA_50: 'N',
  SMA_200: 'O',
  
  // MOMENTUM (P-T)
  RSI: 'P',
  MACD_HIST: 'Q',
  DIVERGENCE: 'R',
  ADX: 'S',
  STOCH_K: 'T',
  
  // VOLATILITY (U-X) - Volatility indicators only
  VOL_REGIME: 'U',      // Volatility regime (ATR/Price)
  BBP_SIGNAL: 'V',      // Bollinger Band Position signal
  ATR: 'W',
  BOLLINGER_BP: 'X',
  
  // TARGET (Y-Z) - Target and R:R Quality
  TARGET: 'Y',
  RR_QUALITY: 'Z',      // NEW: Moved from J (old position)
  
  // LEVELS (AA-AE) - Support/Resistance and position sizing
  SUPPORT: 'AA',
  RESISTANCE: 'AB',
  ATR_STOP: 'AC',
  ATR_TARGET: 'AD',
  POSITION_SIZE: 'AE',  // NEW: Moved from AF to AE
  
  // NOTES (AF) - NOTES column removed, only LAST STATE remains
  LAST_STATE: 'AF'      // Moved from AG to AF after NOTES removal
};

// ============================================================================
// DASHBOARD SHEET COLUMNS (30 columns: A-AD, no R:R Quality)
// ============================================================================

const DASH_COLUMNS = {
  // IDENTITY
  TICKER: 'A',
  
  // SIGNALING (B-D)
  SIGNAL: 'B',
  PATTERNS: 'C',        // NEW: Moved from Z
  DECISION: 'D',
  
  // PRICE / VOLUME (E-H)
  PRICE: 'E',
  CHANGE_PCT: 'F',
  VOL_TREND: 'G',
  VOL_REGIME: 'H',      // NEW: Moved from X
  
  // PERFORMANCE (I-L)
  ATH_TRUE: 'I',
  ATH_DIFF_PCT: 'J',
  ATH_ZONE: 'K',        // NEW: Moved from old I
  FUNDAMENTAL: 'L',     // NEW: Moved from old C, kept under PERFORMANCE
  
  // TREND (M-P) - TREND_SCORE removed
  TREND_STATE: 'M',
  SMA_20: 'N',
  SMA_50: 'O',
  SMA_200: 'P',
  
  // MOMENTUM (Q-V)
  RSI: 'Q',
  MACD_HIST: 'R',
  DIVERGENCE: 'S',
  ADX: 'T',
  STOCH_K: 'U',
  BBP_SIGNAL: 'V',      // NEW: Moved from AA
  
  // VOLATILITY (W-Z)
  ATR: 'W',
  BOLLINGER_BP: 'X',
  TARGET: 'Y',
  SUPPORT: 'Z',         // From CALC AA
  
  // TARGET (AA-AD)
  RESISTANCE: 'AA',     // From CALC AB
  ATR_STOP: 'AB',       // From CALC AC
  ATR_TARGET: 'AC',     // From CALC AD
  POSITION_SIZE: 'AD'   // From CALC AE
  // Note: R:R Quality (CALC Z) is NOT in DASHBOARD
};

// ============================================================================
// COLUMN GROUPS FOR HEADERS
// ============================================================================

const CALC_GROUPS = {
  IDENTITY: { start: 'A', end: 'A', label: 'IDENTITY', color: '#37474F' },
  SIGNALING: { start: 'B', end: 'D', label: 'SIGNALING', color: '#1565C0' },
  PRICE_VOLUME: { start: 'E', end: 'G', label: 'PRICE / VOLUME', color: '#D84315' },
  PERFORMANCE: { start: 'H', end: 'K', label: 'PERFORMANCE', color: '#1976D2' },  // H-K includes FUNDAMENTAL
  TREND: { start: 'L', end: 'O', label: 'TREND', color: '#00838F' },
  MOMENTUM: { start: 'P', end: 'T', label: 'MOMENTUM', color: '#F57C00' },
  VOLATILITY: { start: 'U', end: 'X', label: 'VOLATILITY', color: '#C62828' },
  TARGET: { start: 'Y', end: 'AE', label: 'TARGET', color: '#AD1457' },  // Y-AE single group (includes Target, R:R, Support, Resistance, ATR STOP/TARGET, Position Size)
  NOTES: { start: 'AF', end: 'AF', label: 'NOTES', color: '#616161' }
};

const DASH_GROUPS = {
  IDENTITY: { start: 'A', end: 'A', label: 'IDENTITY', color: '#37474F' },
  SIGNALING: { start: 'B', end: 'D', label: 'SIGNALING', color: '#1565C0' },
  PRICE_VOLUME: { start: 'E', end: 'G', label: 'PRICE / VOLUME', color: '#D84315' },
  PERFORMANCE: { start: 'H', end: 'K', label: 'PERFORMANCE', color: '#1976D2' },  // H-K includes FUNDAMENTAL
  TREND: { start: 'L', end: 'O', label: 'TREND', color: '#00838F' },
  MOMENTUM: { start: 'P', end: 'T', label: 'MOMENTUM', color: '#F57C00' },
  VOLATILITY: { start: 'U', end: 'X', label: 'VOLATILITY', color: '#C62828' },
  TARGET: { start: 'Y', end: 'AE', label: 'TARGET', color: '#AD1457' }  // Y-AE single group (matches CALCULATIONS structure)
};

// ============================================================================
// COLUMN HEADERS (ROW 2)
// ============================================================================

const CALC_HEADERS = [
  'Ticker',           // A
  'SIGNAL',           // B
  'PATTERNS',         // C
  'DECISION',         // D
  'Price',            // E
  'Change %',         // F
  'Vol Trend',        // G - RVOL (only volume indicator)
  'ATH (TRUE)',       // H
  'ATH Diff %',       // I
  'ATH ZONE',         // J
  'FUNDAMENTAL',      // K
  'Trend State',      // L
  'SMA 20',           // M
  'SMA 50',           // N
  'SMA 200',          // O
  'RSI',              // P
  'MACD Hist',        // Q
  'Divergence',       // R
  'ADX (14)',         // S
  'Stoch %K (14)',    // T
  'VOL REGIME',       // U - Volatility indicator
  'BBP SIGNAL',       // V - Volatility-based signal
  'ATR (14)',         // W - Volatility indicator
  'Bollinger %B',     // X - Volatility indicator
  'Target (3:1)',     // Y
  'R:R Quality',      // Z
  'Support',          // AA
  'Resistance',       // AB
  'ATR STOP',         // AC
  'ATR TARGET',       // AD
  'POSITION SIZE',    // AE
  'LAST STATE'        // AF
];

const DASH_HEADERS = [
  'Ticker',           // A
  'SIGNAL',           // B
  'PATTERNS',         // C
  'DECISION',         // D
  'Price',            // E
  'Change %',         // F
  'Vol Trend',        // G
  'VOL REGIME',       // H
  'ATH (TRUE)',       // I
  'ATH Diff %',       // J
  'ATH ZONE',         // K
  'FUNDAMENTAL',      // L
  'Trend State',      // M
  'SMA 20',           // N
  'SMA 50',           // O
  'SMA 200',          // P
  'RSI',              // Q
  'MACD Hist',        // R
  'Divergence',       // S
  'ADX (14)',         // T
  'Stoch %K (14)',    // U
  'BBP SIGNAL',       // V
  'ATR (14)',         // W
  'Bollinger %B',     // X
  'Target (3:1)',     // Y
  'Support',          // Z (from CALC AA)
  'Resistance',       // AA (from CALC AB)
  'ATR STOP',         // AB (from CALC AC)
  'ATR TARGET',       // AC (from CALC AD)
  'POSITION SIZE'     // AD (from CALC AE)
];

// ============================================================================
// MAPPING: CALCULATIONS → DASHBOARD
// ============================================================================

const CALC_TO_DASH_MAPPING = {
  'A': 'A',   // Ticker
  'B': 'B',   // SIGNAL
  'C': 'C',   // PATTERNS
  'D': 'D',   // DECISION
  'E': 'E',   // Price
  'F': 'F',   // Change %
  'G': 'G',   // Vol Trend
  'H': 'H',   // VOL REGIME
  'I': 'I',   // ATH (TRUE)
  'J': 'J',   // ATH Diff %
  'K': 'K',   // ATH ZONE
  'L': 'L',   // FUNDAMENTAL
  'M': 'M',   // Trend State
  'N': 'N',   // SMA 20
  'O': 'O',   // SMA 50
  'P': 'P',   // SMA 200
  'Q': 'Q',   // RSI
  'R': 'R',   // MACD Hist
  'S': 'S',   // Divergence
  'T': 'T',   // ADX (14)
  'U': 'U',   // Stoch %K (14)
  'V': 'V',   // BBP SIGNAL
  'W': 'W',   // ATR (14)
  'X': 'X',   // Bollinger %B
  'Y': 'Y',   // Target (3:1)
  'Z': null,  // R:R Quality - NOT in DASHBOARD
  'AA': 'Z',  // Support (CALC AA → DASH Z)
  'AB': 'AA', // Resistance (CALC AB → DASH AA)
  'AC': 'AB', // ATR STOP (CALC AC → DASH AB)
  'AD': 'AC', // ATR TARGET (CALC AD → DASH AC)
  'AE': 'AD', // POSITION SIZE (CALC AE → DASH AD)
  // AF=LAST_STATE not mapped to Dashboard
};

// ============================================================================
// HELPER FUNCTIONS
// ============================================================================

/**
 * Convert column letter to number (A=1, B=2, ..., Z=26, AA=27, etc.)
 */
function columnLetterToNumber(letter) {
  let column = 0;
  for (let i = 0; i < letter.length; i++) {
    column = column * 26 + (letter.charCodeAt(i) - 64);
  }
  return column;
}

/**
 * Convert column number to letter (1=A, 2=B, ..., 26=Z, 27=AA, etc.)
 */
function columnNumberToLetter(column) {
  let temp, letter = '';
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}

/**
 * Get column number from CALC_COLUMNS object
 */
function getCalcColumnNumber(key) {
  return columnLetterToNumber(CALC_COLUMNS[key]);
}

/**
 * Get column number from DASH_COLUMNS object
 */
function getDashColumnNumber(key) {
  return columnLetterToNumber(DASH_COLUMNS[key]);
}

/**
 * Get Dashboard column letter from Calculations column letter
 */
function mapCalcToDash(calcColumn) {
  return CALC_TO_DASH_MAPPING[calcColumn] || null;
}

// ============================================================================
// EXPORTS
// ============================================================================

// For Google Apps Script (no module system)
if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    CALC_COLUMNS,
    DASH_COLUMNS,
    CALC_GROUPS,
    DASH_GROUPS,
    CALC_HEADERS,
    DASH_HEADERS,
    CALC_TO_DASH_MAPPING,
    columnLetterToNumber,
    columnNumberToLetter,
    getCalcColumnNumber,
    getDashColumnNumber,
    mapCalcToDash
  };
}
