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

  // Row 2: Gold price (I2-L2)
  dashboard.getRange("I2")
    .setValue("GOLD")
    .setBackground("#FFD700").setFontColor("#000000").setFontWeight("bold")
    .setFontSize(10)
    .setHorizontalAlignment("center").setVerticalAlignment("middle");

  dashboard.getRange("J2")
    .setFormula('=AI("What is gold price in india today per gram for 24 carat. Insert only number in rupee format")')
    .setBackground("#FFF8DC").setFontColor("#B8860B").setFontWeight("bold")
    .setFontSize(10)
    .setHorizontalAlignment("center").setVerticalAlignment("middle")
    .setNumberFormat("#,##0.00");

  dashboard.getRange("K2")
    .setFormula('=AI("What is gold price in india yesterday per gram for 24 carat. Insert only number in rupee format")')
    .setBackground("#FFF8DC").setFontColor("#B8860B").setFontWeight("bold")
    .setFontSize(10)
    .setHorizontalAlignment("center").setVerticalAlignment("middle")
    .setNumberFormat("#,##0.00");

  dashboard.getRange("L2")
    .setFormula('=IFERROR((J2-K2)/K2,0)')
    .setBackground("#FFF8DC").setFontWeight("bold")
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
    dashboard.getRange("A3:AH3").breakApart();
  } catch (e) {
    // Ignore if no merges exist
  }

  dashboard.getRange("A3:AI3").clearContent();
  styleGroup("A3:A3", "IDENTITY", "#37474F");        // Dark Blue-Grey (A)
  styleGroup("B3:F3", "SIGNALING", "#1565C0");       // Blue (B-F: MARKET RATING, DECISION, SIGNAL, PATTERNS, CONSENSUS PRICE)
  styleGroup("G3:I3", "PRICE / VOLUME", "#D84315");  // Deep Orange (G-I: Price, Change%, RVOL)
  styleGroup("J3:P3", "PERFORMANCE", "#1976D2");     // Medium Blue (J-P: ATH Diff%, 52WH Diff%, 52WL Diff%, P/E, EPS, ATH ZONE, FUNDAMENTAL)
  styleGroup("Q3:S3", "TREND", "#00838F");           // Cyan (Q-S: SMA 20%/50%/200%)
  styleGroup("T3:X3", "MOMENTUM", "#F57C00");        // Orange (T-X: RSI, MACD, Div, ADX, Stoch)
  styleGroup("Y3:AB3", "VOLATILITY", "#C62828");     // Red (Y-AB: VOLATILITY REGIME, BBP SIGNAL, ATR, Bollinger %B)
  styleGroup("AC3:AI3", "TARGET", "#AD1457");        // Pink (AC-AI: All target-related)
  dashboard.getRange("A3:AI3").setWrap(true);

  // Row 4: Column headers (moved from old Row 3) - 35 columns A-AI
  const headers = [[
    "Ticker",              // A
    "MARKET RATING",       // B
    "DECISION",            // C
    "SIGNAL",              // D
    "PATTERNS",            // E
    "CONSENSUS PRICE",     // F
    "Price",               // G
    "Change %",            // H
    "RVOL",                // I (Relative Volume)
    "ATH Diff %",          // J (NEW - from CALCULATIONS K)
    "52WH Diff %",         // K (NEW - calculated)
    "52WL Diff %",         // L (NEW - calculated)
    "P/E",                 // M (NEW - from DATA)
    "EPS",                 // N (NEW - from DATA)
    "ATH ZONE",            // O (from CALCULATIONS L)
    "FUNDAMENTAL",         // P (from CALCULATIONS M)
    "SMA 20 %",            // Q (calculated % from CALCULATIONS O)
    "SMA 50 %",            // R (calculated % from CALCULATIONS P)
    "SMA 200 %",           // S (calculated % from CALCULATIONS Q)
    "RSI",                 // T (from CALCULATIONS R)
    "MACD Hist",           // U (from CALCULATIONS S)
    "Divergence",          // V (from CALCULATIONS T)
    "ADX (14)",            // W (from CALCULATIONS U)
    "Stoch %K (14)",       // X (from CALCULATIONS V)
    "VOLATILITY REGIME",   // Y (from CALCULATIONS W)
    "BBP SIGNAL",          // Z (from CALCULATIONS X)
    "ATR (14)",            // AA (from CALCULATIONS Y)
    "Bollinger %B",        // AB (from CALCULATIONS Z)
    "Target (3:1)",        // AC (from CALCULATIONS AA)
    "R:R",                 // AD (from CALCULATIONS AB)
    "Support",             // AE (from CALCULATIONS AC)
    "Resistance",          // AF (from CALCULATIONS AD)
    "ATR STOP",            // AG (from CALCULATIONS AE)
    "ATR TARGET",          // AH (from CALCULATIONS AF)
    "POSITION SIZE"        // AI (from CALCULATIONS AG) (35 columns total)
  ]];

  dashboard.getRange(4, 1, 1, 35)
    .setValues(headers)
    .setBackground("#0D0D0D").setFontColor("#FFD700").setFontWeight("bold")
    .setHorizontalAlignment("center").setVerticalAlignment("middle")
    .setWrap(true);

  // Freeze panes
  dashboard.setFrozenRows(4);
  dashboard.setFrozenColumns(1);

  // Borders for rows 3-4 only (preserve row 1-2 borders from above)
  dashboard.getRange("A3:AI4")
    .setBorder(true, true, true, true, true, true, "#FFFFFF", SpreadsheetApp.BorderStyle.SOLID);

  // Sentinel note
  dashboard.getRange("A1").setNote(SENTINEL);
  
  // Setup dropdowns for control row
  setupControlRowDropdowns();
}

function refreshDashboardData(dashboard, ss, DATA_START_ROW, preserveCheckboxes = false) {
  // Get locale separator
  const SEP = (/^(en|en_)/.test(ss.getSpreadsheetLocale())) ? "," : ";";
  
  // Clear existing data (35 columns A-AI)
  dashboard.getRange(DATA_START_ROW, 1, 1000, 35).clearContent();

  // Only reset checkboxes on initial setup (not when filters change)
  if (!preserveCheckboxes) {
    dashboard.getRange("B1").setValue(false);
    dashboard.getRange("D1").setValue(false);
    dashboard.getRange("H1").setValue(false);
    dashboard.getRange("J1").setValue(false);
    dashboard.getRange("L1").setValue(false);
    dashboard.getRange("N1").setValue(false);
  }

  // Get filter settings
  const usaSelected = dashboard.getRange("B1").getValue() === true;
  const indiaSelected = dashboard.getRange("D1").getValue() === true;
  const categoryFilter = String(dashboard.getRange("F1").getValue() || "ALL").toUpperCase().trim();
  const sortColumnName = dashboard.getRange("B2").getValue() || "Change %";
  
  // Get source data
  const calc = ss.getSheetByName("CALCULATIONS");
  const data = ss.getSheetByName("DATA");
  const input = ss.getSheetByName("INPUT");
  
  if (!calc || !data || !input) {
    ss.toast("Required sheets not found", "Error", 3);
    return;
  }
  
  // Read all data from CALCULATIONS (starting row 3)
  const calcData = calc.getRange(3, 1, calc.getLastRow() - 2, 34).getValues();
  
  // Read DATA row 2 (52WH, 52WL) and row 3 (P/E, EPS)
  const dataRow2 = data.getRange(2, 1, 1, data.getLastColumn()).getValues()[0];
  const dataRow3 = data.getRange(3, 1, 1, data.getLastColumn()).getValues()[0];
  
  // Read INPUT for filtering
  const inputData = input.getRange(3, 1, input.getLastRow() - 2, 3).getValues(); // A, B, C columns
  
  // Build lookup maps
  const inputMap = {};
  for (let i = 0; i < inputData.length; i++) {
    const ticker = String(inputData[i][0] || "").trim().toUpperCase();
    if (ticker) {
      inputMap[ticker] = {
        country: String(inputData[i][1] || "").trim().toUpperCase(),
        category: String(inputData[i][2] || "").trim().toUpperCase()
      };
    }
  }
  
  // Build DATA lookup (find ticker positions in DATA row 2)
  const dataMap = {};
  for (let i = 0; i < dataRow2.length; i += 7) {
    const ticker = String(dataRow2[i] || "").trim().toUpperCase();
    if (ticker) {
      dataMap[ticker] = {
        wh52: dataRow2[i + 2] || 0,  // 52WH at offset +2
        wl52: dataRow2[i + 4] || 0,  // 52WL at offset +4
        pe: dataRow3[i + 3] || 0,    // P/E at offset +3 in row 3
        eps: dataRow3[i + 5] || 0    // EPS at offset +5 in row 3
      };
    }
  }
  
  // Process and filter data
  const outputData = [];
  
  for (let i = 0; i < calcData.length; i++) {
    const row = calcData[i];
    const ticker = String(row[0] || "").trim().toUpperCase();
    
    if (!ticker) continue;
    
    // Apply filters
    const inputInfo = inputMap[ticker];
    if (!inputInfo) continue;
    
    // Country filter
    if (usaSelected && indiaSelected) {
      // Both selected - show all
    } else if (usaSelected && inputInfo.country !== "USA") {
      continue;
    } else if (indiaSelected && inputInfo.country !== "INDIA") {
      continue;
    } else if (!usaSelected && !indiaSelected) {
      // Neither selected - show nothing
      continue;
    }
    
    // Category filter
    if (categoryFilter && categoryFilter !== "ALL") {
      const categories = categoryFilter.split(",").map(c => c.trim());
      const inputCategory = inputInfo.category.replace(/\s+/g, "");
      let matchFound = false;
      for (let c of categories) {
        const cleanCat = c.replace(/\s+/g, "");
        if (inputCategory.includes(cleanCat)) {
          matchFound = true;
          break;
        }
      }
      if (!matchFound) continue;
    }
    
    // Get DATA values
    const dataInfo = dataMap[ticker] || {wh52: 0, wl52: 0, pe: 0, eps: 0};
    const price = Number(row[6]) || 0; // Price is column G (index 6)
    
    // Calculate diff percentages
    const athDiff = Number(row[10]) || 0; // ATH Diff % from CALCULATIONS K (index 10)
    const wh52Diff = dataInfo.wh52 > 0 ? (price - dataInfo.wh52) / dataInfo.wh52 : 0;
    const wl52Diff = dataInfo.wl52 > 0 ? (price - dataInfo.wl52) / dataInfo.wl52 : 0;
    
    // Calculate SMA percentages (if price is above SMA, it's positive; if below, it's negative)
    const sma20 = Number(row[14]) || 0; // SMA 20 from CALCULATIONS O (index 14)
    const sma50 = Number(row[15]) || 0; // SMA 50 from CALCULATIONS P (index 15)
    const sma200 = Number(row[16]) || 0; // SMA 200 from CALCULATIONS Q (index 16)
    
    const sma20Pct = sma20 > 0 ? (price - sma20) / sma20 : 0;
    const sma50Pct = sma50 > 0 ? (price - sma50) / sma50 : 0;
    const sma200Pct = sma200 > 0 ? (price - sma200) / sma200 : 0;
    
    // Build output row (35 columns A-AI)
    const outRow = [
      row[0],   // A: Ticker
      row[1],   // B: MARKET RATING
      row[2],   // C: DECISION
      row[3],   // D: SIGNAL
      row[4],   // E: PATTERNS
      row[5],   // F: CONSENSUS PRICE
      row[6],   // G: Price
      row[7],   // H: Change %
      row[8],   // I: RVOL
      athDiff,  // J: ATH Diff % (from CALCULATIONS K)
      wh52Diff, // K: 52WH Diff % (calculated)
      wl52Diff, // L: 52WL Diff % (calculated)
      dataInfo.pe,  // M: P/E (from DATA)
      dataInfo.eps, // N: EPS (from DATA)
      row[11],  // O: ATH ZONE (from CALCULATIONS L)
      row[12],  // P: FUNDAMENTAL (from CALCULATIONS M)
      sma20Pct, // Q: SMA 20 % (calculated)
      sma50Pct, // R: SMA 50 % (calculated)
      sma200Pct,// S: SMA 200 % (calculated)
      row[17],  // T: RSI (from CALCULATIONS R)
      row[18],  // U: MACD Hist (from CALCULATIONS S)
      row[19],  // V: Divergence (from CALCULATIONS T)
      row[20],  // W: ADX (from CALCULATIONS U)
      row[21],  // X: Stoch %K (from CALCULATIONS V)
      row[22],  // Y: VOLATILITY REGIME (from CALCULATIONS W)
      row[23],  // Z: BBP SIGNAL (from CALCULATIONS X)
      row[24],  // AA: ATR (from CALCULATIONS Y)
      row[25],  // AB: Bollinger %B (from CALCULATIONS Z)
      row[26],  // AC: Target (from CALCULATIONS AA)
      row[27],  // AD: R:R (from CALCULATIONS AB)
      row[28],  // AE: Support (from CALCULATIONS AC)
      row[29],  // AF: Resistance (from CALCULATIONS AD)
      row[30],  // AG: ATR STOP (from CALCULATIONS AE)
      row[31],  // AH: ATR TARGET (from CALCULATIONS AF)
      row[32]   // AI: POSITION SIZE (from CALCULATIONS AG)
    ];
    
    outputData.push(outRow);
  }
  
  // Sort data
  const columnMap = {
    "Ticker": 0, "MARKET RATING": 1, "DECISION": 2, "SIGNAL": 3, "PATTERNS": 4,
    "CONSENSUS PRICE": 5, "Price": 6, "Change %": 7, "RVOL": 8,
    "ATH Diff %": 9, "52WH Diff %": 10, "52WL Diff %": 11, "P/E": 12, "EPS": 13,
    "ATH ZONE": 14, "FUNDAMENTAL": 15, "SMA 20 %": 16, "SMA 50 %": 17, "SMA 200 %": 18,
    "RSI": 19, "MACD Hist": 20, "Divergence": 21, "ADX (14)": 22, "Stoch %K (14)": 23,
    "VOLATILITY REGIME": 24, "BBP SIGNAL": 25, "ATR (14)": 26, "Bollinger %B": 27,
    "Target (3:1)": 28, "R:R": 29, "Support": 30, "Resistance": 31,
    "ATR STOP": 32, "ATR TARGET": 33, "POSITION SIZE": 34
  };
  
  const sortColIndex = columnMap[sortColumnName] || 7; // Default to Change %
  
  // Determine sort order based on column name
  // For ATH%, 52WH%, 52WL% - sort ASCENDING (big negative to high positive)
  // For all other columns - sort DESCENDING (highest to lowest)
  const ascendingColumns = ['ATH Diff %', '52WH Diff %', '52WL Diff %'];
  const isAscending = ascendingColumns.includes(sortColumnName);
  
  outputData.sort((a, b) => {
    const aVal = Number(a[sortColIndex]) || 0;
    const bVal = Number(b[sortColIndex]) || 0;
    return isAscending ? (aVal - bVal) : (bVal - aVal); // Ascending or Descending
  });
  
  // Write data to DASHBOARD
  if (outputData.length > 0) {
    dashboard.getRange(DATA_START_ROW, 1, outputData.length, 35).setValues(outputData);
  }
  
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
  const TOTAL_COLS = 35;

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
  const dataRange = sh.getRange(DATA_START_ROW, 1, numRows, 35);

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

  // Number formats (updated for 35 columns A-AI)
  sh.getRange(DATA_START_ROW, 6, numRows, 1).setNumberFormat("#,##0.00");  // F: CONSENSUS PRICE
  sh.getRange(DATA_START_ROW, 7, numRows, 1).setNumberFormat("#,##0.00");  // G: Price
  sh.getRange(DATA_START_ROW, 8, numRows, 1).setNumberFormat("0.00%");     // H: Change%
  sh.getRange(DATA_START_ROW, 9, numRows, 1).setNumberFormat("0.00");      // I: RVOL
  sh.getRange(DATA_START_ROW, 10, numRows, 3).setNumberFormat("0.00%");    // J-L: ATH Diff%, 52WH Diff%, 52WL Diff%
  sh.getRange(DATA_START_ROW, 13, numRows, 1).setNumberFormat("0.00");     // M: P/E
  sh.getRange(DATA_START_ROW, 14, numRows, 1).setNumberFormat("0.00");     // N: EPS
  sh.getRange(DATA_START_ROW, 15, numRows, 1).setNumberFormat("@");        // O: ATH ZONE
  sh.getRange(DATA_START_ROW, 16, numRows, 1).setNumberFormat("@");        // P: FUNDAMENTAL
  sh.getRange(DATA_START_ROW, 17, numRows, 3).setNumberFormat("0.00%");    // Q-S: SMA %
  sh.getRange(DATA_START_ROW, 20, numRows, 1).setNumberFormat("0.0");      // T: RSI
  sh.getRange(DATA_START_ROW, 21, numRows, 1).setNumberFormat("0.000");    // U: MACD
  sh.getRange(DATA_START_ROW, 23, numRows, 1).setNumberFormat("0.0");      // W: ADX
  sh.getRange(DATA_START_ROW, 24, numRows, 1).setNumberFormat("0.00%");    // X: Stoch
  sh.getRange(DATA_START_ROW, 25, numRows, 1).setNumberFormat("@");        // Y: VOLATILITY REGIME
  sh.getRange(DATA_START_ROW, 26, numRows, 1).setNumberFormat("@");        // Z: BBP SIGNAL
  sh.getRange(DATA_START_ROW, 27, numRows, 1).setNumberFormat("#,##0.00"); // AA: ATR
  sh.getRange(DATA_START_ROW, 28, numRows, 1).setNumberFormat("0.00");     // AB: Bollinger %B
  sh.getRange(DATA_START_ROW, 29, numRows, 1).setNumberFormat("#,##0.00"); // AC: Target (3:1)
  sh.getRange(DATA_START_ROW, 30, numRows, 1).setNumberFormat("0.00");     // AD: R:R
  sh.getRange(DATA_START_ROW, 31, numRows, 1).setNumberFormat("#,##0.00"); // AE: Support
  sh.getRange(DATA_START_ROW, 32, numRows, 1).setNumberFormat("#,##0.00"); // AF: Resistance
  sh.getRange(DATA_START_ROW, 33, numRows, 1).setNumberFormat("#,##0.00"); // AG: ATR STOP
  sh.getRange(DATA_START_ROW, 34, numRows, 1).setNumberFormat("#,##0.00"); // AH: ATR TARGET
  sh.getRange(DATA_START_ROW, 35, numRows, 1).setNumberFormat("@");        // AI: POSITION SIZE

  // Apply conditional formatting rules
  applyConditionalFormatting(sh, DATA_START_ROW, numRows, C_GREEN, C_RED, C_BLUE);
}

function applyDashboardBloombergFormatting_(sh, DATA_START_ROW) {
  if (!sh) return;

  const C_BLUE = "#E3F2FD";   // Light blue (default background)
  const C_GREEN = "#C8E6C9";  // Light green (positive)
  const C_RED = "#FFCDD2";    // Light red (negative)
  const HEADER_DARK = "#1F1F1F";
  const TOTAL_COLS = 35; // Updated to 35 columns (A-AI)

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
    // No columns to hide - all 35 columns are visible
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
  const dataRange = sh.getRange(DATA_START_ROW, 1, numRows, 35); // Updated to 35 columns

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

  // Column widths (35 columns)
  for (let c = 1; c <= 35; c++) sh.setColumnWidth(c, 85);

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

  // Number formats (updated for 35 columns A-AI)
  sh.getRange(DATA_START_ROW, 6, numRows, 1).setNumberFormat("#,##0.00");  // F: CONSENSUS PRICE
  sh.getRange(DATA_START_ROW, 7, numRows, 1).setNumberFormat("#,##0.00");  // G: Price
  sh.getRange(DATA_START_ROW, 8, numRows, 1).setNumberFormat("0.00%");     // H: Change%
  sh.getRange(DATA_START_ROW, 9, numRows, 1).setNumberFormat("0.00");      // I: RVOL
  sh.getRange(DATA_START_ROW, 10, numRows, 3).setNumberFormat("0.00%");    // J-L: ATH Diff%, 52WH Diff%, 52WL Diff%
  sh.getRange(DATA_START_ROW, 13, numRows, 1).setNumberFormat("0.00");     // M: P/E
  sh.getRange(DATA_START_ROW, 14, numRows, 1).setNumberFormat("0.00");     // N: EPS
  sh.getRange(DATA_START_ROW, 15, numRows, 1).setNumberFormat("@");        // O: ATH ZONE
  sh.getRange(DATA_START_ROW, 16, numRows, 1).setNumberFormat("@");        // P: FUNDAMENTAL
  sh.getRange(DATA_START_ROW, 17, numRows, 3).setNumberFormat("0.00%");    // Q-S: SMA %
  sh.getRange(DATA_START_ROW, 20, numRows, 1).setNumberFormat("0.0");      // T: RSI
  sh.getRange(DATA_START_ROW, 21, numRows, 1).setNumberFormat("0.000");    // U: MACD
  sh.getRange(DATA_START_ROW, 23, numRows, 1).setNumberFormat("0.0");      // W: ADX
  sh.getRange(DATA_START_ROW, 24, numRows, 1).setNumberFormat("0.00%");    // X: Stoch
  sh.getRange(DATA_START_ROW, 25, numRows, 1).setNumberFormat("@");        // Y: VOLATILITY REGIME
  sh.getRange(DATA_START_ROW, 26, numRows, 1).setNumberFormat("@");        // Z: BBP SIGNAL
  sh.getRange(DATA_START_ROW, 27, numRows, 1).setNumberFormat("#,##0.00"); // AA: ATR
  sh.getRange(DATA_START_ROW, 28, numRows, 1).setNumberFormat("0.00");     // AB: Bollinger %B
  sh.getRange(DATA_START_ROW, 29, numRows, 1).setNumberFormat("#,##0.00"); // AC: Target
  sh.getRange(DATA_START_ROW, 30, numRows, 1).setNumberFormat("0.00");     // AD: R:R
  sh.getRange(DATA_START_ROW, 31, numRows, 1).setNumberFormat("#,##0.00"); // AE: Support
  sh.getRange(DATA_START_ROW, 32, numRows, 1).setNumberFormat("#,##0.00"); // AF: Resistance
  sh.getRange(DATA_START_ROW, 33, numRows, 1).setNumberFormat("#,##0.00"); // AG: ATR STOP
  sh.getRange(DATA_START_ROW, 34, numRows, 1).setNumberFormat("#,##0.00"); // AH: ATR TARGET
  sh.getRange(DATA_START_ROW, 35, numRows, 1).setNumberFormat("@");        // AI: POSITION SIZE (3:1)
  sh.getRange(DATA_START_ROW, 30, numRows, 1).setNumberFormat("0.00");     // AD: R:R Quality
  sh.getRange(DATA_START_ROW, 31, numRows, 1).setNumberFormat("#,##0.00"); // AE: Support
  sh.getRange(DATA_START_ROW, 32, numRows, 1).setNumberFormat("#,##0.00"); // AF: Resistance
  sh.getRange(DATA_START_ROW, 33, numRows, 1).setNumberFormat("#,##0.00"); // AG: ATR STOP
  sh.getRange(DATA_START_ROW, 34, numRows, 1).setNumberFormat("@");        // AH: POSITION SIZE
  sh.getRange(DATA_START_ROW, 34, numRows, 1).setNumberFormat("#,##0.00"); // AH: ATR TARGET
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

  // RVOL (I) - Green for high volume, Red for low
  add(`=$I${r0}>=1.5`, C_GREEN, 9);
  add(`=$I${r0}<=0.85`, C_RED, 9);

  // ATH Diff % (J) - Red when < -5%, Green when >= -5%
  add(`=$J${r0}<-0.05`, C_RED, 10);
  add(`=$J${r0}>=-0.05`, C_GREEN, 10);

  // 52WH Diff % (K) - Red when < -5%, Green when >= -5%
  add(`=$K${r0}<-0.05`, C_RED, 11);
  add(`=$K${r0}>=-0.05`, C_GREEN, 11);

  // 52WL Diff % (L) - Green far above 52WL (>= 20%), Red near 52WL (<= 5%)
  add(`=$L${r0}>=0.20`, C_GREEN, 12);
  add(`=$L${r0}<=0.05`, C_RED, 12);

  // P/E (M) - No conditional formatting (just number format)
  // EPS (N) - No conditional formatting (just number format)

  // ATH ZONE (O) - Green at/near ATH, Red in correction
  add(`=REGEXMATCH($O${r0},"AT ATH|NEAR ATH")`, C_GREEN, 15);
  add(`=REGEXMATCH($O${r0},"DEEP VALUE|CORRECTION")`, C_RED, 15);

  // FUNDAMENTAL (P) - Green for value, Red for expensive
  add(`=$P${r0}="VALUE"`, C_GREEN, 16);
  add(`=REGEXMATCH($P${r0},"EXPENSIVE|PRICED FOR PERFECTION|ZOMBIE")`, C_RED, 16);

  // SMA % (Q/R/S) - Green when positive (price above SMA), Red when negative (price below SMA)
  add(`=$Q${r0}>0`, C_GREEN, 17);
  add(`=$Q${r0}<0`, C_RED, 17);
  add(`=$R${r0}>0`, C_GREEN, 18);
  add(`=$R${r0}<0`, C_RED, 18);
  add(`=$S${r0}>0`, C_GREEN, 19);
  add(`=$S${r0}<0`, C_RED, 19);

  // RSI (T) - Green oversold (opportunity), Red overbought
  add(`=$T${r0}<=30`, C_GREEN, 20);
  add(`=$T${r0}>=70`, C_RED, 20);

  // MACD Hist (U) - Green positive, Red negative
  add(`=$U${r0}>0`, C_GREEN, 21);
  add(`=$U${r0}<0`, C_RED, 21);

  // Divergence (V) - Green bullish, Red bearish
  add(`=REGEXMATCH($V${r0},"BULL")`, C_GREEN, 22);
  add(`=REGEXMATCH($V${r0},"BEAR")`, C_RED, 22);

  // ADX (W) - Green strong trend, no red (weak trend stays blue)
  add(`=$W${r0}>=25`, C_GREEN, 23);

  // Stoch %K (X) - Green oversold, Red overbought
  add(`=$X${r0}<=0.2`, C_GREEN, 24);
  add(`=$X${r0}>=0.8`, C_RED, 24);

  // VOL REGIME (Y) - Green low vol, Red extreme vol
  add(`=$Y${r0}="LOW VOL"`, C_GREEN, 25);
  add(`=$Y${r0}="EXTREME VOL"`, C_RED, 25);

  // BBP SIGNAL (Z) - Green oversold/mean reversion, Red overbought
  add(`=REGEXMATCH($Z${r0},"EXTREME OVERSOLD|MEAN REVERSION")`, C_GREEN, 26);
  add(`=REGEXMATCH($Z${r0},"EXTREME OVERBOUGHT")`, C_RED, 26);

  // ATR (AA) - Green low volatility, Red high volatility
  add(`=IFERROR($AA${r0}/$G${r0},0)<=0.02`, C_GREEN, 27);
  add(`=IFERROR($AA${r0}/$G${r0},0)>=0.05`, C_RED, 27);

  // Bollinger %B (AB) - Green oversold, Red overbought
  add(`=$AB${r0}<=0.2`, C_GREEN, 28);
  add(`=$AB${r0}>=0.8`, C_RED, 28);

  // Target (AC) - Green good upside, Red limited upside
  add(`=AND($AC${r0}>0,$AC${r0}>=$G${r0}*1.05)`, C_GREEN, 29);
  add(`=AND($AC${r0}>0,$AC${r0}<=$G${r0}*1.01)`, C_RED, 29);

  // R:R Quality (AD) - Green good R:R, Red poor R:R
  add(`=$AD${r0}>=3`, C_GREEN, 30);
  add(`=$AD${r0}<=1`, C_RED, 30);

  // Support (AE) - Green at/near support, Red below support
  add(`=AND($AE${r0}>0,$G${r0}>=$AE${r0},$G${r0}<=$AE${r0}*1.01)`, C_GREEN, 31);
  add(`=AND($AE${r0}>0,$G${r0}<$AE${r0})`, C_RED, 31);

  // Resistance (AF) - Green far from resistance, Red at resistance
  add(`=AND($AF${r0}>0,$G${r0}<=$AF${r0}*0.90)`, C_GREEN, 32);
  add(`=AND($AF${r0}>0,$G${r0}>=$AF${r0}*0.995)`, C_RED, 32);

  // ATR STOP (AG) - No conditional formatting
  // POSITION SIZE (AH) - No conditional formatting

  sh.setConditionalFormatRules(rules);
}

/**
 * Apply conditional formatting to market index cells (E2, H2, and L2)
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
  
  // Gold price % change (L2) - Red font for negative, keep golden background
  indexRules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThan(0)
      .setFontColor("#FF0000")  // Red font for negative
      .setRanges([sh.getRange("L2")])
      .build()
  );
  
  indexRules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThanOrEqualTo(0)
      .setFontColor("#B8860B")  // Golden font for positive/zero
      .setRanges([sh.getRange("L2")])
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
    TARGET: "#AD1457"           // Pink (includes all target-related columns AA-AI)
  };

  const FG = "#FFFFFF";

  // GROUPS array for 35 columns (A-AI) - Matches new structure
  const GROUPS = [
    { name: "IDENTITY", c1: 1, c2: 1, color: COLORS.IDENTITY },           // A
    { name: "SIGNALING", c1: 2, c2: 6, color: COLORS.SIGNALING },         // B-F (MARKET RATING, DECISION, SIGNAL, PATTERNS, CONSENSUS PRICE)
    { name: "PRICE / VOLUME", c1: 7, c2: 9, color: COLORS.PRICE_VOLUME }, // G-I (Price, Change%, RVOL)
    { name: "PERFORMANCE", c1: 10, c2: 16, color: COLORS.PERFORMANCE },   // J-P (ATH Diff%, 52WH Diff%, 52WL Diff%, P/E, EPS, ATH ZONE, FUNDAMENTAL)
    { name: "TREND", c1: 17, c2: 19, color: COLORS.TREND },               // Q-S (SMA 20%/50%/200%)
    { name: "MOMENTUM", c1: 20, c2: 24, color: COLORS.MOMENTUM },         // T-X (RSI, MACD, Div, ADX, Stoch)
    { name: "VOLATILITY", c1: 25, c2: 28, color: COLORS.VOLATILITY },     // Y-AB (VOLATILITY REGIME, BBP SIGNAL, ATR, Bollinger %B)
    { name: "TARGET", c1: 29, c2: 35, color: COLORS.TARGET }              // AC-AI (Target, R:R, Support, Res, ATR STOP, ATR TARGET, Position)
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
    sh.getRange(3, 1, 1, 35).breakApart();
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
 * 
 * Task 4.1: When H1 checkbox changes (INVEST/TRADE mode toggle), 
 * this function now triggers automatic data repopulation to refresh
 * the dashboard with the new mode settings.
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
    
    // Refresh dashboard data to apply mode change (preserve checkbox states)
    // Task 4.1: Automatic data repopulation when H1 changes
    const DATA_START_ROW = 5;
    refreshDashboardData(dashboard, ss, DATA_START_ROW, true);
    
    // Show notification
    const mode = modeValue ? 'INVEST' : 'TRADE';
    ss.toast(`Switched to ${mode} mode`, '✓ Mode Updated', 3);
    
  } finally {
    props.deleteProperty('SYNCING_MODE');
  }
}

/**
 * Refresh dashboard data without rebuilding layout (wrapper for J1 checkbox)
 * 
 * Task 4.1: When J1 checkbox changes (Dashboard refresh), this function
 * triggers automatic data repopulation and resets the checkbox to false.
 */
function refreshDashboardDataFromCheckbox() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dashboard = ss.getSheetByName('DASHBOARD');
  
  if (!dashboard) return;
  
  ss.toast('Refreshing dashboard data...', '⚙️ Processing', -1);
  
  try {
    const DATA_START_ROW = 5;
    // Preserve checkbox states when manually refreshing
    // Task 4.1: Automatic data repopulation when J1 changes
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
    // Task 4.1: Reset J1 checkbox to false after refresh completes
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
    const headers = dashboard.getRange('A4:AI4').getValues()[0];
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
    
    const dataRange = dashboard.getRange(5, 1, lastRow - 4, 35);
    
    // Determine sort order based on column name
    // For ATH%, 52WH%, 52WL% - sort ASCENDING (big negative to high positive)
    // For all other columns - sort DESCENDING (highest to lowest)
    const ascendingColumns = ['ATH Diff %', '52WH Diff %', '52WL Diff %'];
    const isAscending = ascendingColumns.includes(columnName);
    
    // Sort with appropriate order
    dataRange.sort({
      column: sortColumnNumber,
      ascending: isAscending
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
  const headers = dashboard.getRange('B4:AI4').getValues()[0]
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
