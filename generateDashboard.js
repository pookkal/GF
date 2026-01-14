/**
 * generateDashboard.js
 * Optimized version of generateDashboardSheet with progressive loading and error handling
 * Preserves exact filtering logic and Bloomberg formatting from Code.js
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
    const tickers = getCleanTickers(input);
    
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
    .setBackground("#004D40").setFontColor("white").setFontWeight("bold")
    .setHorizontalAlignment("center").setVerticalAlignment("middle");

  dashboard.getRange("I1")
    .setFormula(`=IFERROR(GOOGLEFINANCE("INDEXSP:.INX"${SEP}"price")${SEP}0)`)
    .setBackground("#004D40").setFontColor("white").setFontWeight("bold")
    .setHorizontalAlignment("center").setVerticalAlignment("middle")
    .setNumberFormat("#,##0.00");

  dashboard.getRange("J1")
    .setFormula(`=IFERROR(GOOGLEFINANCE("INDEXSP:.INX"${SEP}"changepct")/100${SEP}0)`)
    .setBackground("#004D40").setFontColor("white").setFontWeight("bold")
    .setHorizontalAlignment("center").setVerticalAlignment("middle")
    .setNumberFormat("0.00%");

  // Row 2 group headers
  const styleGroup = (a1, label, bg) => {
    dashboard.getRange(a1).merge()
      .setValue(label)
      .setBackground(bg).setFontColor("white").setFontWeight("bold")
      .setHorizontalAlignment("center").setVerticalAlignment("middle");
  };

  dashboard.getRange("A2:AF2").clearContent();
  styleGroup("A2:A2", "IDENTITY", "#263238");
  styleGroup("B2:D2", "SIGNALING", "#0D47A1");
  styleGroup("E2:F2", "PRICE", "#1B5E20");
  styleGroup("G2:J2", "PERFORMANCE", "#004D40");
  styleGroup("K2:O2", "TREND", "#2E7D32");
  styleGroup("P2:T2", "MOMENTUM", "#33691E");
  styleGroup("U2:Z2", "VOLUME / VOLATILITY", "#B71C1C");
  styleGroup("AA2:AF2", "TARGET", "#6A1B9A");
  dashboard.getRange("A2:AF2").setWrap(true);

  // Row 3 column headers
  const headers = [[
    "Ticker", "SIGNAL", "FUNDAMENTAL", "DECISION", "Price", "Change %",
    "ATH (TRUE)", "ATH Diff %", "ATH ZONE", "R:R Quality", "Trend Score", "Trend State",
    "SMA 20", "SMA 50", "SMA 200",
    "RSI", "MACD Hist", "Divergence", "ADX (14)", "Stoch %K (14)",
    "Vol Trend", "ATR (14)", "Bollinger %B", "VOL REGIME", "POSITION SIZE", "EVENT",
    "BBP SIGNAL", "Support", "Resistance", "Target (3:1)", "ATR STOP", "ATR TARGET"
  ]];

  dashboard.getRange(3, 1, 1, 32)
    .setValues(headers)
    .setBackground("#111111").setFontColor("white").setFontWeight("bold")
    .setHorizontalAlignment("center").setVerticalAlignment("middle")
    .setWrap(true);

  // Freeze panes
  dashboard.setFrozenRows(3);
  dashboard.setFrozenColumns(1);

  // White border for top header rows
  dashboard.getRange("A1:AF3")
    .setBorder(true, true, true, true, true, true, "#FFFFFF", SpreadsheetApp.BorderStyle.SOLID);

  // Sentinel note
  dashboard.getRange("A1").setNote(SENTINEL);
}

function refreshDashboardData(dashboard, ss, DATA_START_ROW) {
  // Clear existing data
  dashboard.getRange(DATA_START_ROW, 1, 1000, 35).clearContent();

  // Reset checkboxes to false
  dashboard.getRange("B1").setValue(false);
  dashboard.getRange("D1").setValue(false);

  // Get locale separator
  const SEP = (/^(en|en_)/.test(ss.getSpreadsheetLocale())) ? "," : ";";

  // Filter formula - pulls columns from CALCULATIONS in new order
  const filterFormula =
    '=IFERROR(' +
    'SORT(' +
    'FILTER({' +
    'CALCULATIONS!$A$3:$A,' +  // A: Ticker
    'CALCULATIONS!$B$3:$B,' +  // B: SIGNAL
    'CALCULATIONS!$C$3:$C,' +  // C: FUNDAMENTAL
    'CALCULATIONS!$D$3:$D,' +  // D: DECISION
    'CALCULATIONS!$E$3:$E,' +  // E: Price
    'CALCULATIONS!$F$3:$F,' +  // F: Change %
    'CALCULATIONS!$H$3:$H,' +  // G: ATH (TRUE)
    'CALCULATIONS!$I$3:$I,' +  // H: ATH Diff %
    'CALCULATIONS!$AD$3:$AD,' + // I: ATH ZONE
    'CALCULATIONS!$J$3:$J,' +  // J: R:R Quality
    'CALCULATIONS!$K$3:$K,' +  // K: Trend Score
    'CALCULATIONS!$L$3:$L,' +  // L: Trend State
    'CALCULATIONS!$M$3:$M,' +  // M: SMA 20
    'CALCULATIONS!$N$3:$N,' +  // N: SMA 50
    'CALCULATIONS!$O$3:$O,' +  // O: SMA 200
    'CALCULATIONS!$P$3:$P,' +  // P: RSI
    'CALCULATIONS!$Q$3:$Q,' +  // Q: MACD Hist
    'CALCULATIONS!$R$3:$R,' +  // R: Divergence
    'CALCULATIONS!$S$3:$S,' +  // S: ADX (14)
    'CALCULATIONS!$T$3:$T,' +  // T: Stoch %K (14)
    'CALCULATIONS!$G$3:$G,' +  // U: Vol Trend
    'CALCULATIONS!$X$3:$X,' +  // V: ATR (14)
    'CALCULATIONS!$Y$3:$Y,' +  // W: Bollinger %B
    'CALCULATIONS!$AC$3:$AC,' + // X: VOL REGIME
    'CALCULATIONS!$Z$3:$Z,' +  // Y: POSITION SIZE
    'CALCULATIONS!$AF$3:$AF,' + // Z: EVENT (was PATTERNS)
    'CALCULATIONS!$AE$3:$AE,' + // AA: BBP SIGNAL
    'CALCULATIONS!$U$3:$U,' +  // AB: Support
    'CALCULATIONS!$V$3:$V,' +  // AC: Resistance
    'CALCULATIONS!$W$3:$W,' +  // AD: Target (3:1)
    'CALCULATIONS!$AG$3:$AG,' + // AE: ATR STOP
    'CALCULATIONS!$AH$3:$AH' + // AF: ATR TARGET
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
  applyDashboardGroupMapAndColors_(dashboard);
}

function applyDashboardBloombergFormatting_(sh, DATA_START_ROW) {
  if (!sh) return;

  const C_WHITE = "#FFFFFF";
  const C_GREEN = "#C6EFCE";
  const C_RED = "#FFC7CE";
  const C_GREY = "#E7E6E6";
  const HEADER_DARK = "#1F1F1F";
  const TOTAL_COLS = 32;

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
  const dataRange = sh.getRange(DATA_START_ROW, 1, numRows, 32);

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

  // Column widths
  for (let c = 1; c <= 32; c++) sh.setColumnWidth(c, 85);

  // Row heights
  sh.setRowHeight(1, 22);
  sh.setRowHeight(2, 18);
  sh.setRowHeight(3, 22);
  sh.setRowHeights(DATA_START_ROW, numRows, 54);

  // Data range styling
  dataRange
    .setBackground(C_WHITE)
    .setHorizontalAlignment("left")
    .setVerticalAlignment("middle")
    .setWrap(true)
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);

  dataRange.setBorder(true, true, true, true, true, true, "#000000", SpreadsheetApp.BorderStyle.SOLID);

  sh.getRange(1, 1, 3, TOTAL_COLS)
    .setBorder(true, true, true, true, true, true, "#FFFFFF", SpreadsheetApp.BorderStyle.SOLID);

  // Number formats
  sh.getRange(DATA_START_ROW, 5, numRows, 1).setNumberFormat("#,##0.00");  // E: Price
  sh.getRange(DATA_START_ROW, 6, numRows, 1).setNumberFormat("0.00%");     // F: Change%
  sh.getRange(DATA_START_ROW, 7, numRows, 1).setNumberFormat("#,##0.00");  // G: ATH
  sh.getRange(DATA_START_ROW, 8, numRows, 1).setNumberFormat("0.00%");     // H: ATH Diff%
  sh.getRange(DATA_START_ROW, 10, numRows, 1).setNumberFormat("0.00");     // J: R:R
  sh.getRange(DATA_START_ROW, 13, numRows, 3).setNumberFormat("#,##0.00"); // M-O: SMAs
  sh.getRange(DATA_START_ROW, 16, numRows, 1).setNumberFormat("0.0");      // P: RSI
  sh.getRange(DATA_START_ROW, 17, numRows, 1).setNumberFormat("0.000");    // Q: MACD
  sh.getRange(DATA_START_ROW, 19, numRows, 1).setNumberFormat("0.0");      // S: ADX
  sh.getRange(DATA_START_ROW, 20, numRows, 1).setNumberFormat("0.00%");    // T: Stoch
  sh.getRange(DATA_START_ROW, 21, numRows, 1).setNumberFormat("0.00");     // U: Vol Trend (RVOL)
  sh.getRange(DATA_START_ROW, 22, numRows, 1).setNumberFormat("#,##0.00"); // V: ATR
  sh.getRange(DATA_START_ROW, 23, numRows, 1).setNumberFormat("0.00");     // W: Bollinger %B
  sh.getRange(DATA_START_ROW, 28, numRows, 4).setNumberFormat("#,##0.00"); // AB-AE: Support/Res/Target/ATR STOP
  sh.getRange(DATA_START_ROW, 32, numRows, 1).setNumberFormat("#,##0.00"); // AF: ATR TARGET

  // Conditional formatting for market indices in row 1
  const indexRules = [];
  
  // NIFTY 50 % change (G1)
  indexRules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied("=$G$1<0")
      .setBackground(C_RED)
      .setRanges([sh.getRange("G1")])
      .build()
  );
  
  // S&P 500 % change (J1)
  indexRules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied("=$J$1<0")
      .setBackground(C_RED)
      .setRanges([sh.getRange("J1")])
      .build()
  );
  
  sh.setConditionalFormatRules(indexRules.concat(sh.getConditionalFormatRules()));

  // Apply conditional formatting rules
  applyConditionalFormatting(sh, DATA_START_ROW, numRows, C_GREEN, C_RED, C_GREY);

  safeHideNotes_();
  clearTailFormats_(lastDataRow);
}

function applyConditionalFormatting(sh, r0, numRows, C_GREEN, C_RED, C_GREY) {
  const rules = [];
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

  // SIGNAL (B)
  add(`=REGEXMATCH($B${r0},"Breakout|Trend Continuation|Mean Reversion")`, C_GREEN, 2);
  add(`=REGEXMATCH($B${r0},"Stop-Out|Risk-Off")`, C_RED, 2);
  add(`=REGEXMATCH($B${r0},"Volatility Squeeze|Range-Bound|Hold")`, C_GREY, 2);

  // FUNDAMENTAL (C)
  add(`=$C${r0}="VALUE"`, C_GREEN, 3);
  add(`=$C${r0}="FAIR"`, C_GREY, 3);
  add(`=REGEXMATCH($C${r0},"EXPENSIVE|PRICED FOR PERFECTION|ZOMBIE")`, C_RED, 3);

  // DECISION (D)
  add(`=REGEXMATCH($D${r0},"Trade Long|Accumulate|Add in Dip")`, C_GREEN, 4);
  add(`=REGEXMATCH($D${r0},"Stop-Out|Avoid|Reduce|Take Profit")`, C_RED, 4);
  add(`=REGEXMATCH($D${r0},"Hold|Monitor|LOADING")`, C_GREY, 4);

  // PRICE (E) and Change% (F)
  add(`=$F${r0}>0`, C_GREEN, 5);
  add(`=$F${r0}<0`, C_RED, 5);
  add(`=OR($F${r0}=0,$F${r0}="")`, C_GREY, 5);
  add(`=$F${r0}>0`, C_GREEN, 6);
  add(`=$F${r0}<0`, C_RED, 6);
  add(`=OR($F${r0}=0,$F${r0}="")`, C_GREY, 6);

  // ATH (G) / ATH Diff % (H)
  add(`=AND($G${r0}>0,$E${r0}>=$G${r0}*0.995)`, C_GREEN, 7);
  add(`=AND($G${r0}>0,$E${r0}<=$G${r0}*0.80)`, C_RED, 7);
  add(`=AND($G${r0}>0,$E${r0}>$G${r0}*0.80,$E${r0}<$G${r0}*0.995)`, C_GREY, 7);
  add(`=$H${r0}>=-0.05`, C_GREEN, 8);
  add(`=$H${r0}<=-0.20`, C_RED, 8);
  add(`=AND($H${r0}>-0.20,$H${r0}<-0.05)`, C_GREY, 8);

  // R:R (J)
  add(`=$J${r0}>=3`, C_GREEN, 10);
  add(`=$J${r0}<1.5`, C_RED, 10);
  add(`=AND($J${r0}>=1.5,$J${r0}<3)`, C_GREY, 10);

  // Trend Score (K)
  add(`=LEN($K${r0})>=3`, C_GREEN, 11);
  add(`=LEN($K${r0})<=1`, C_RED, 11);
  add(`=LEN($K${r0})=2`, C_GREY, 11);

  // Trend State (L)
  add(`=$L${r0}="BULL"`, C_GREEN, 12);
  add(`=$L${r0}="BEAR"`, C_RED, 12);
  add(`=AND($L${r0}<>"BULL",$L${r0}<>"BEAR")`, C_GREY, 12);

  // SMAs (M/N/O)
  add(`=AND($M${r0}>0,$E${r0}>=$M${r0})`, C_GREEN, 13);
  add(`=AND($M${r0}>0,$E${r0}<$M${r0})`, C_RED, 13);
  add(`=AND($N${r0}>0,$E${r0}>=$N${r0})`, C_GREEN, 14);
  add(`=AND($N${r0}>0,$E${r0}<$N${r0})`, C_RED, 14);
  add(`=AND($O${r0}>0,$E${r0}>=$O${r0})`, C_GREEN, 15);
  add(`=AND($O${r0}>0,$E${r0}<$O${r0})`, C_RED, 15);

  // RSI (P)
  add(`=$P${r0}<=30`, C_GREEN, 16);
  add(`=$P${r0}>=70`, C_RED, 16);
  add(`=AND($P${r0}>30,$P${r0}<70)`, C_GREY, 16);

  // MACD Hist (Q)
  add(`=$Q${r0}>0`, C_GREEN, 17);
  add(`=$Q${r0}<0`, C_RED, 17);
  add(`=OR($Q${r0}=0,$Q${r0}="")`, C_GREY, 17);

  // Divergence (R)
  add(`=REGEXMATCH($R${r0},"BULL")`, C_GREEN, 18);
  add(`=REGEXMATCH($R${r0},"BEAR")`, C_RED, 18);
  add(`=OR($R${r0}="—",$R${r0}="",NOT(REGEXMATCH($R${r0},"BULL|BEAR")))`, C_GREY, 18);

  // ADX (S)
  add(`=$S${r0}>=25`, C_GREEN, 19);
  add(`=$S${r0}<15`, C_GREY, 19);
  add(`=AND($S${r0}>=15,$S${r0}<25)`, C_GREY, 19);

  // Stoch %K (T)
  add(`=$T${r0}<=0.2`, C_GREEN, 20);
  add(`=$T${r0}>=0.8`, C_RED, 20);
  add(`=AND($T${r0}>0.2,$T${r0}<0.8)`, C_GREY, 20);

  // Vol Trend RVOL (U)
  add(`=$U${r0}>=1.5`, C_GREEN, 21);
  add(`=$U${r0}<=0.85`, C_RED, 21);
  add(`=AND($U${r0}>0.85,$U${r0}<1.5)`, C_GREY, 21);

  // ATR (V)
  add(`=IFERROR($V${r0}/$E${r0},0)<=0.02`, C_GREEN, 22);
  add(`=IFERROR($V${r0}/$E${r0},0)>=0.05`, C_RED, 22);
  add(`=AND(IFERROR($V${r0}/$E${r0},0)>0.02,IFERROR($V${r0}/$E${r0},0)<0.05)`, C_GREY, 22);

  // Bollinger %B (W)
  add(`=$W${r0}<=0.2`, C_GREEN, 23);
  add(`=$W${r0}>=0.8`, C_RED, 23);
  add(`=AND($W${r0}>0.2,$W${r0}<0.8)`, C_GREY, 23);

  // Support (AB)
  add(`=AND($AB${r0}>0,$E${r0}<$AB${r0})`, C_RED, 28);
  add(`=AND($AB${r0}>0,$E${r0}>=$AB${r0},$E${r0}<=$AB${r0}*1.01)`, C_GREEN, 28);
  add(`=AND($AB${r0}>0,$E${r0}>$AB${r0}*1.01)`, C_GREY, 28);

  // Resistance (AC)
  add(`=AND($AC${r0}>0,$E${r0}>=$AC${r0}*0.995)`, C_RED, 29);
  add(`=AND($AC${r0}>0,$E${r0}<=$AC${r0}*0.90)`, C_GREEN, 29);
  add(`=AND($AC${r0}>0,$E${r0}>$AC${r0}*0.90,$E${r0}<$AC${r0}*0.995)`, C_GREY, 29);

  // Target (AD)
  add(`=AND($AD${r0}>0,$AD${r0}>=$E${r0}*1.05)`, C_GREEN, 30);
  add(`=AND($AD${r0}>0,$AD${r0}<=$E${r0}*1.01)`, C_RED, 30);
  add(`=AND($AD${r0}>0,$AD${r0}>$E${r0}*1.01,$AD${r0}<$E${r0}*1.05)`, C_GREY, 30);

  sh.setConditionalFormatRules(rules);
}

function applyDashboardGroupMapAndColors_(sh) {
  if (!sh) return;

  const COLORS = {
    SIGNAL: "#1F4FD8",
    PRICE: "#0F766E",
    PERF: "#374151",
    TREND: "#14532D",
    MOM: "#7C2D12",
    VOL: "#B71C1C",
    PATTERNS: "#6A1B9A",
    INST: "#4A148C",
    NOTES: "#111827"
  };

  const FG = "#FFFFFF";

  const GROUPS = [
    { name: "SIGNALING", c1: 2, c2: 4, color: COLORS.SIGNAL },
    { name: "PRICE", c1: 5, c2: 6, color: COLORS.PRICE },
    { name: "PERFORMANCE", c1: 7, c2: 10, color: COLORS.PERF },
    { name: "TREND", c1: 11, c2: 15, color: COLORS.TREND },
    { name: "MOMENTUM", c1: 16, c2: 20, color: COLORS.MOM },
    { name: "VOLUME / VOLATILITY", c1: 21, c2: 26, color: COLORS.VOL },
    { name: "TARGET", c1: 27, c2: 32, color: COLORS.PATTERNS }
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

  GROUPS.forEach(g => {
    style(2, g.c1, g.c2, g.color);
    const r2 = sh.getRange(2, g.c1, 1, g.c2 - g.c1 + 1);
    try { r2.breakApart(); } catch (e) { }
    if (g.c1 !== g.c2) r2.merge();
    r2.setValue(g.name);
    style(3, g.c1, g.c2, g.color);
  });
}
