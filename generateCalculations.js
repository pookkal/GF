/**
* ==============================================================================
* STABLE_MASTER_ALL_CLEAN_v3.1_KIRO_OPTIMIZED
* ==============================================================================
*/

// Delay constants for staggered formula writing (in milliseconds)
// These delays prevent calculation engine overload, especially on Android app
const DELAY_AFTER_MAIN_FORMULAS = 12500;  // 12.5 seconds - allows calculation engine to process bulk formulas (columns E-AF)
const DELAY_AFTER_CD_FORMULAS = 2000;     // 2 seconds - shorter delay for smaller formula set (columns C-D)

// Column headers for CALCULATIONS sheet (34 columns: A-AH)
const CALC_HEADERS = [
  'Ticker',           // A
  'MARKET RATING',    // B (NEW - references INPUT D)
  'DECISION',         // C (old B formula)
  'SIGNAL',           // D (old C formula)
  'PATTERNS',         // E (old D formula - GETPATTERNS)
  'CONSENSUS PRICE',  // F (NEW - references INPUT E)
  'Price',            // G (old E formula)
  'Change %',         // H (shifted from F)
  'Vol Trend',        // I (shifted from G)
  'ATH (TRUE)',       // J (shifted from H)
  'ATH Diff %',       // K (shifted from I)
  'ATH ZONE',         // L (shifted from J)
  'FUNDAMENTAL',      // M (shifted from K)
  'Trend State',      // N (shifted from L)
  'SMA 20',           // O (shifted from M)
  'SMA 50',           // P (shifted from N)
  'SMA 200',          // Q (shifted from O)
  'RSI',              // R (shifted from P)
  'MACD Hist',        // S (shifted from Q)
  'Divergence',       // T (shifted from R)
  'ADX (14)',         // U (shifted from S)
  'Stoch %K (14)',    // V (shifted from T)
  'VOL REGIME',       // W (shifted from U)
  'BBP SIGNAL',       // X (shifted from V)
  'ATR (14)',         // Y (shifted from W)
  'Bollinger %B',     // Z (shifted from X)
  'Target (3:1)',     // AA (shifted from Y)
  'R:R Quality',      // AB (shifted from Z)
  'Support',          // AC (shifted from AA)
  'Resistance',       // AD (shifted from AB)
  'ATR STOP',         // AE (shifted from AC)
  'ATR TARGET',       // AF (shifted from AD)
  'POSITION SIZE',    // AG (shifted from AE)
  'LAST STATE'        // AH (shifted from AF)
];

function generateCalculationsSheet() {
  const startTime = new Date();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  try {
    // Validate required sheets
    const dataSheet = ss.getSheetByName("DATA");
    const inputSheet = ss.getSheetByName("INPUT");
    
    if (!dataSheet || !inputSheet) {
      ss.toast('Required sheets (DATA or INPUT) not found.', '❌ Error', 3);
      return;
    }

    const tickers = getCleanTickers(inputSheet);
    if (!tickers || tickers.length === 0) {
      ss.toast('No tickers found in INPUT sheet.', '⚠️ Warning', 3);
      return;
    }

    let calc = ss.getSheetByName("CALCULATIONS") || ss.insertSheet("CALCULATIONS");

    // Locale separator: US typically ","; EU typically ";"
    const locale = (ss.getSpreadsheetLocale() || "").toLowerCase();
    const SEP = (/^(en|en_)/.test(locale)) ? "," : ";";

    // Clear existing content
    calc.clear().clearFormats();

    // Ensure sheet has enough columns (34 total: A-AH)
    const maxCols = calc.getMaxColumns();
    if (maxCols < 34) {
      calc.insertColumnsAfter(maxCols, 34 - maxCols);
    }

    // PHASE 1: Setup headers
    Logger.log('Setting up headers...');
    setupHeaders(calc, ss, SEP);
    
    // PHASE 2: Write tickers (progressive)
    Logger.log(`Writing ${tickers.length} tickers...`);
    writeTickers(calc, tickers);
    
    SpreadsheetApp.flush();
    
    // PHASE 3: Write formulas sequentially
    Logger.log('Writing formulas sequentially...');
    writeFormulas(calc, tickers, SEP);
    
    const elapsed = ((new Date() - startTime) / 1000).toFixed(2);
    Logger.log(`CALCULATIONS sheet completed in ${elapsed}s`);
    ss.toast(`✓ CALCULATIONS sheet generated in ${elapsed}s`, 'Success', 3);
    
  } catch (error) {
    Logger.log(`Error in generateCalculationsSheet: ${error.stack}`);
    ss.toast(`Failed to generate CALCULATIONS: ${error.message}`, '❌ Error', 5);
  }
}

function setupHeaders(calc, ss, SEP) {
  // ROW 1: GROUP HEADERS (MERGED) + timestamp
  const syncTime = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "yyyy-MM-dd HH:mm:ss");
  
  const styleGroup = (a1, label, bg) => {
    calc.getRange(a1).merge()
      .setValue(label)
      .setBackground(bg)
      .setFontColor("white")
      .setFontWeight("bold")
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle")
      .setBorder(true, true, true, true, false, false, "white", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  };

  // Define group colors
  const COLORS = {
    IDENTITY: "#37474F",
    SIGNALING: "#1565C0",
    PRICE_VOLUME: "#D84315",
    PERFORMANCE: "#1976D2",  // PERFORMANCE includes FUNDAMENTAL (H-K)
    TREND: "#00838F",
    MOMENTUM: "#F57C00",
    VOLATILITY: "#C62828",
    TARGET: "#AD1457",  // TARGET includes old LEVELS (Y-AE)
    NOTES: "#616161"
  };

  // ROW 1: Group headers with CORRECT merges
  styleGroup("A1:A1", "IDENTITY", COLORS.IDENTITY);
  styleGroup("B1:F1", "SIGNALING", COLORS.SIGNALING);  // B-F: MARKET RATING (INPUT D), DECISION, SIGNAL, PATTERNS (GETPATTERNS), CONSENSUS PRICE (INPUT E)
  styleGroup("G1:I1", "PRICE / VOLUME", COLORS.PRICE_VOLUME);  // G-I: Price, Change%, Vol Trend
  styleGroup("J1:M1", "PERFORMANCE", COLORS.PERFORMANCE);  // Shifted from H1:K1, includes FUNDAMENTAL
  styleGroup("N1:Q1", "TREND", COLORS.TREND);  // Shifted from L1:O1
  styleGroup("R1:V1", "MOMENTUM", COLORS.MOMENTUM);  // Shifted from P1:T1
  styleGroup("W1:Z1", "VOLATILITY", COLORS.VOLATILITY);  // Shifted from U1:X1
  styleGroup("AA1:AG1", "TARGET", COLORS.TARGET);  // Shifted from Y1:AE1
  styleGroup("AH1:AH1", "NOTES", COLORS.NOTES);  // Shifted from AF1

  // Timestamp in AH1
  calc.getRange("AH1")
    .setValue(syncTime)
    .setBackground("#000000")
    .setFontColor("#00FF00")
    .setFontWeight("bold")
    .setFontSize(9)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");

  // ROW 2: COLUMN HEADERS with matching group colors
  const headerColors = [
    COLORS.IDENTITY,      // A: Ticker
    COLORS.SIGNALING, COLORS.SIGNALING, COLORS.SIGNALING, COLORS.SIGNALING, COLORS.SIGNALING,  // B-F: MARKET RATING (INPUT D), DECISION, SIGNAL, PATTERNS (GETPATTERNS), CONSENSUS PRICE (INPUT E)
    COLORS.PRICE_VOLUME, COLORS.PRICE_VOLUME, COLORS.PRICE_VOLUME,  // G-I: Price, Change%, Vol Trend
    COLORS.PERFORMANCE, COLORS.PERFORMANCE, COLORS.PERFORMANCE, COLORS.PERFORMANCE,  // J-M: ATH TRUE, ATH Diff%, ATH ZONE, FUNDAMENTAL
    COLORS.TREND, COLORS.TREND, COLORS.TREND, COLORS.TREND,  // N-Q: Trend State, SMA 20/50/200
    COLORS.MOMENTUM, COLORS.MOMENTUM, COLORS.MOMENTUM, COLORS.MOMENTUM, COLORS.MOMENTUM,  // R-V: RSI, MACD, Div, ADX, Stoch
    COLORS.VOLATILITY, COLORS.VOLATILITY, COLORS.VOLATILITY, COLORS.VOLATILITY,  // W-Z: VOL REGIME, BBP SIGNAL, ATR, Bollinger %B
    COLORS.TARGET, COLORS.TARGET,  // AA-AB: Target, R:R Quality
    COLORS.TARGET, COLORS.TARGET, COLORS.TARGET, COLORS.TARGET, COLORS.TARGET,  // AC-AG: Support, Resistance, ATR STOP/TARGET, Position Size
    COLORS.NOTES  // AH: LAST STATE
  ];

  // Set Row 2 headers with group colors
  calc.getRange(2, 1, 1, 34)
    .setValues([CALC_HEADERS])
    .setFontColor("white")
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setWrap(true)
    .setBorder(true, true, true, true, true, true, "white", SpreadsheetApp.BorderStyle.SOLID);
  
  // Apply individual colors to Row 2
  for (let i = 0; i < headerColors.length; i++) {
    calc.getRange(2, i + 1).setBackground(headerColors[i]);
  }

  // Set column widths (Bloomberg style)
  calc.setColumnWidth(1, 80);   // A: Ticker
  calc.setColumnWidth(2, 120);  // B: MARKET RATING (NEW - from INPUT D)
  calc.setColumnWidth(3, 100);  // C: DECISION (old B formula)
  calc.setColumnWidth(4, 80);   // D: SIGNAL (old C formula)
  calc.setColumnWidth(5, 150);  // E: PATTERNS (old D formula - GETPATTERNS)
  calc.setColumnWidth(6, 110);  // F: CONSENSUS PRICE (NEW - from INPUT E)
  calc.setColumnWidth(7, 80);   // G: Price (old E formula)
  calc.setColumnWidth(8, 80);   // H: Change % (shifted from F)
  calc.setColumnWidth(9, 80);   // I: Vol Trend (shifted from G)
  calc.setColumnWidth(10, 90);  // J: ATH (TRUE) (shifted from H)
  calc.setColumnWidth(11, 80);  // K: ATH Diff % (shifted from I)
  calc.setColumnWidth(12, 120); // L: ATH ZONE (shifted from J)
  calc.setColumnWidth(13, 140); // M: FUNDAMENTAL (shifted from K)
  calc.setColumnWidth(14, 100); // N: Trend State (shifted from L)
  calc.setColumnWidth(15, 80);  // O: SMA 20 (shifted from M)
  calc.setColumnWidth(16, 80);  // P: SMA 50 (shifted from N)
  calc.setColumnWidth(17, 80);  // Q: SMA 200 (shifted from O)
  calc.setColumnWidth(18, 70);  // R: RSI (shifted from P)
  calc.setColumnWidth(19, 80);  // S: MACD Hist (shifted from Q)
  calc.setColumnWidth(20, 100); // T: Divergence (shifted from R)
  calc.setColumnWidth(21, 70);  // U: ADX (shifted from S)
  calc.setColumnWidth(22, 90);  // V: Stoch %K (shifted from T)
  calc.setColumnWidth(23, 110); // W: VOL REGIME (shifted from U)
  calc.setColumnWidth(24, 130); // X: BBP SIGNAL (shifted from V)
  calc.setColumnWidth(25, 70);  // Y: ATR (shifted from W)
  calc.setColumnWidth(26, 100); // Z: Bollinger %B (shifted from X)
  calc.setColumnWidth(27, 80);  // AA: Target (shifted from Y)
  calc.setColumnWidth(28, 90);  // AB: R:R Quality (shifted from Z)
  calc.setColumnWidth(29, 80);  // AC: Support (shifted from AA)
  calc.setColumnWidth(30, 90);  // AD: Resistance (shifted from AB)
  calc.setColumnWidth(31, 90);  // AE: ATR STOP (shifted from AC)
  calc.setColumnWidth(32, 100); // AF: ATR TARGET (shifted from AD)
  calc.setColumnWidth(33, 120); // AG: POSITION SIZE (shifted from AE)
  calc.setColumnWidth(34, 120); // AH: LAST STATE (shifted from AF)

  // Set row heights
  calc.setRowHeight(1, 30);  // Row 1: Group headers
  calc.setRowHeight(2, 40);  // Row 2: Column headers
}

/**
 * Apply Bloomberg-style formatting to data rows
 * - Alternating row colors (lighter versions of group colors)
 * - Left alignment for text columns, right for numeric
 * - Borders for all cells
 * @param {Sheet} calc - CALCULATIONS sheet
 * @param {number} numRows - Number of data rows (tickers)
 */
function applyBloombergDataFormatting(calc, numRows) {
  if (numRows === 0) return;
  
  const startRow = 3; // Data starts at row 3
  const numCols = 34; // A-AH
  
  // Define lighter versions of group colors (30% lighter)
  const LIGHT_COLORS = {
    IDENTITY: "#78909C",      // Lighter version of #37474F
    SIGNALING: "#64B5F6",     // Lighter version of #1565C0
    PRICE_VOLUME: "#FF8A65",  // Lighter version of #D84315
    PERFORMANCE: "#64B5F6",   // Lighter version of #1976D2
    TREND: "#4DD0E1",         // Lighter version of #00838F
    MOMENTUM: "#FFB74D",      // Lighter version of #F57C00
    VOLATILITY: "#EF5350",    // Lighter version of #C62828
    TARGET: "#EC407A",        // Lighter version of #AD1457
    NOTES: "#9E9E9E"          // Lighter version of "#616161"
  };
  
  // Column background colors (matching group colors)
  const columnColors = [
    LIGHT_COLORS.IDENTITY,      // A: Ticker
    LIGHT_COLORS.SIGNALING, LIGHT_COLORS.SIGNALING, LIGHT_COLORS.SIGNALING, LIGHT_COLORS.SIGNALING, LIGHT_COLORS.SIGNALING,  // B-F
    LIGHT_COLORS.PRICE_VOLUME, LIGHT_COLORS.PRICE_VOLUME, LIGHT_COLORS.PRICE_VOLUME,  // G-I
    LIGHT_COLORS.PERFORMANCE, LIGHT_COLORS.PERFORMANCE, LIGHT_COLORS.PERFORMANCE, LIGHT_COLORS.PERFORMANCE,  // J-M
    LIGHT_COLORS.TREND, LIGHT_COLORS.TREND, LIGHT_COLORS.TREND, LIGHT_COLORS.TREND,  // N-Q
    LIGHT_COLORS.MOMENTUM, LIGHT_COLORS.MOMENTUM, LIGHT_COLORS.MOMENTUM, LIGHT_COLORS.MOMENTUM, LIGHT_COLORS.MOMENTUM,  // R-V
    LIGHT_COLORS.VOLATILITY, LIGHT_COLORS.VOLATILITY, LIGHT_COLORS.VOLATILITY, LIGHT_COLORS.VOLATILITY,  // W-Z
    LIGHT_COLORS.TARGET, LIGHT_COLORS.TARGET,  // AA-AB
    LIGHT_COLORS.TARGET, LIGHT_COLORS.TARGET, LIGHT_COLORS.TARGET, LIGHT_COLORS.TARGET, LIGHT_COLORS.TARGET,  // AC-AG
    LIGHT_COLORS.NOTES  // AH
  ];
  
  // Apply formatting to all data rows
  const dataRange = calc.getRange(startRow, 1, numRows, numCols);
  
  // Apply borders to all cells - black borders for professional look
  dataRange.setBorder(
    true, true, true, true, true, true,  // top, left, bottom, right, vertical, horizontal
    "#000000",  // Black borders
    SpreadsheetApp.BorderStyle.SOLID
  );
  
  // Apply background colors column by column
  for (let col = 0; col < numCols; col++) {
    const colRange = calc.getRange(startRow, col + 1, numRows, 1);
    colRange.setBackground(columnColors[col]);
  }
  
  // Set text alignment - ALL data cells to left alignment
  dataRange.setHorizontalAlignment("left");
  
  // Set font color to black for better readability
  dataRange.setFontColor("#000000");
  
  // Set font size
  dataRange.setFontSize(10);
  
  Logger.log(`Bloomberg-style formatting applied to ${numRows} data rows`);
}

function writeTickers(calc, tickers) {
  if (tickers.length > 0) {
    calc.getRange(3, 1, tickers.length, 1).setValues(tickers.map(t => [t]));
  }
  SpreadsheetApp.flush();
}

function writeFormulas(calc, tickers, SEP) {
  const BLOCK = 7; // DATA block width (must match generateDataSheet)
  
  // Check if long-term signal mode is enabled
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dashboard = ss.getSheetByName('DASHBOARD');
  const useLongTermSignal = dashboard ? dashboard.getRange('H1').getValue() === true : false;
  
  // Get DATA sheet reference for pattern detection
  const dataSheet = ss.getSheetByName('DATA');
  
  // Initialize error tracking array
  const errors = [];
  
  // Track start time for performance measurement
  const startTime = new Date();
  
  Logger.log(`Starting optimized formula generation for ${tickers.length} tickers (batch processing)`);
  
  // Display start notification
  ss.toast(`Processing ${tickers.length} tickers...`, '⏳ Starting', 3);
  
  // STEP 1: Generate all formulas and detect patterns for all tickers
  Logger.log('Step 1: Generating formulas for all tickers...');
  const allFormulas = [];
  
  for (let i = 0; i < tickers.length; i++) {
    const ticker = tickers[i];
    const row = i + 3;
    
    // Pattern detection (non-fatal)
    try {
      const priceData = getPriceDataForTicker(dataSheet, ticker, i, BLOCK);
      
      if (!priceData || priceData.length === 0) {
        Logger.log(`${ticker}: No price data available`);
        setCachedPattern(ticker, '');
      } else {
        const patterns = detectPatterns(priceData, {minBars: 100, minConfidence: 60});
        const patternString = formatPatternsForSheet(patterns);
        setCachedPattern(ticker, patternString);
        Logger.log(`${ticker}: Cached ${patterns.length} patterns`);
      }
    } catch (patternError) {
      Logger.log(`${ticker}: Error detecting patterns - ${patternError.message}`);
      setCachedPattern(ticker, '');
    }
    
    // Generate formulas
    try {
      const formulas = generateTickerFormulas(ticker, row, i, BLOCK, SEP, useLongTermSignal);
      allFormulas.push({ticker, row, formulas});
    } catch (formulaError) {
      Logger.log(`${ticker}: Error generating formulas - ${formulaError.message}`);
      errors.push({ticker, error: formulaError.message, phase: 'formula'});
      allFormulas.push(null); // Placeholder to maintain array alignment
    }
    
    // Display progress
    if ((i + 1) % 10 === 0 || i === tickers.length - 1) {
      const percentage = Math.round(((i + 1) / tickers.length) * 100);
      Logger.log(`Formula generation: ${i + 1}/${tickers.length} (${percentage}%)`);
    }
  }
  
  // STEP 2: Write Phase 1 formulas (columns G-AH) for ALL tickers at once
  Logger.log('Step 2: Writing Phase 1 formulas (columns G-AH) for all tickers...');
  ss.toast('Writing main formulas (G-AH)...', '⏳ Phase 1', 3);
  
  const phase1Data = [];
  for (let i = 0; i < allFormulas.length; i++) {
    if (allFormulas[i] && allFormulas[i].formulas) {
      const sliced = allFormulas[i].formulas.slice(5); // Indices 5-32 (28 columns: G-AH)
      
      // Validate slice length
      if (sliced.length !== 28) {
        Logger.log(`WARNING: ${allFormulas[i].ticker} Phase 1 has ${sliced.length} elements, expected 28. Formula array length: ${allFormulas[i].formulas.length}`);
        // Pad or trim to exactly 28 elements
        while (sliced.length < 28) sliced.push('');
        if (sliced.length > 28) sliced.length = 28;
      }
      
      phase1Data.push(sliced);
    } else {
      phase1Data.push(new Array(28).fill('')); // Empty row for failed formulas
    }
  }
  
  try {
    if (phase1Data.length > 0) {
      // Final validation before writing
      for (let i = 0; i < phase1Data.length; i++) {
        if (phase1Data[i].length !== 28) {
          throw new Error(`Row ${i} has ${phase1Data[i].length} columns, expected 28`);
        }
      }
      
      calc.getRange(3, 7, phase1Data.length, 28).setFormulas(phase1Data);
      Logger.log(`Phase 1 complete: Wrote formulas for columns G-AH (${phase1Data.length} tickers)`);
    }
  } catch (writeError) {
    Logger.log(`Error writing Phase 1 formulas: ${writeError.message}`);
    errors.push({ticker: 'ALL', error: writeError.message, phase: 'write-phase1'});
  }
  
  SpreadsheetApp.flush();
  
  // DELAY 1: Allow calculation engine to process main formulas
  Logger.log(`Applying ${DELAY_AFTER_MAIN_FORMULAS}ms delay after Phase 1...`);
  ss.toast('Waiting for calculations to complete...', '⏳ Delay', 3);
  Utilities.sleep(DELAY_AFTER_MAIN_FORMULAS);
  
  // STEP 3: Write Phase 2 formulas (columns C-F) for ALL tickers at once
  Logger.log('Step 3: Writing Phase 2 formulas (columns C-F) for all tickers...');
  ss.toast('Writing signal and pattern formulas (C-F)...', '⏳ Phase 2', 3);
  
  const phase2Data = [];
  for (let i = 0; i < allFormulas.length; i++) {
    if (allFormulas[i] && allFormulas[i].formulas) {
      const sliced = allFormulas[i].formulas.slice(1, 5); // Indices 1-4 (4 columns: C-F)
      
      // Validate slice length
      if (sliced.length !== 4) {
        Logger.log(`WARNING: ${allFormulas[i].ticker} Phase 2 has ${sliced.length} elements, expected 4. Formula array length: ${allFormulas[i].formulas.length}`);
        // Pad or trim to exactly 4 elements
        while (sliced.length < 4) sliced.push('');
        if (sliced.length > 4) sliced.length = 4;
      }
      
      phase2Data.push(sliced);
    } else {
      phase2Data.push(new Array(4).fill('')); // Empty row for failed formulas
    }
  }
  
  try {
    if (phase2Data.length > 0) {
      // Final validation before writing
      for (let i = 0; i < phase2Data.length; i++) {
        if (phase2Data[i].length !== 4) {
          throw new Error(`Row ${i} has ${phase2Data[i].length} columns, expected 4`);
        }
      }
      
      calc.getRange(3, 3, phase2Data.length, 4).setFormulas(phase2Data);
      Logger.log(`Phase 2 complete: Wrote formulas for columns C-F (${phase2Data.length} tickers)`);
    }
  } catch (writeError) {
    Logger.log(`Error writing Phase 2 formulas: ${writeError.message}`);
    errors.push({ticker: 'ALL', error: writeError.message, phase: 'write-phase2'});
  }
  
  SpreadsheetApp.flush();
  
  // DELAY 2: Allow calculation engine to process C-D formulas
  Logger.log(`Applying ${DELAY_AFTER_CD_FORMULAS}ms delay after Phase 2...`);
  Utilities.sleep(DELAY_AFTER_CD_FORMULAS);
  
  // STEP 4: Write Phase 3 formulas (column B) for ALL tickers at once
  Logger.log('Step 4: Writing Phase 3 formulas (column B) for all tickers...');
  ss.toast('Writing market rating formulas (B)...', '⏳ Phase 3', 3);
  
  const phase3Data = [];
  for (let i = 0; i < allFormulas.length; i++) {
    if (allFormulas[i] && allFormulas[i].formulas) {
      const formula = allFormulas[i].formulas[0]; // Index 0 (MARKET RATING)
      
      // Validate formula exists
      if (typeof formula !== 'string') {
        Logger.log(`WARNING: ${allFormulas[i].ticker} Phase 3 formula is not a string: ${typeof formula}`);
        phase3Data.push(['']);
      } else {
        phase3Data.push([formula]);
      }
    } else {
      phase3Data.push(['']); // Empty row for failed formulas
    }
  }
  
  try {
    if (phase3Data.length > 0) {
      // Final validation before writing
      for (let i = 0; i < phase3Data.length; i++) {
        if (phase3Data[i].length !== 1) {
          throw new Error(`Row ${i} has ${phase3Data[i].length} columns, expected 1`);
        }
      }
      
      calc.getRange(3, 2, phase3Data.length, 1).setFormulas(phase3Data);
      Logger.log(`Phase 3 complete: Wrote formulas for column B (${phase3Data.length} tickers)`);
    }
  } catch (writeError) {
    Logger.log(`Error writing Phase 3 formulas: ${writeError.message}`);
    errors.push({ticker: 'ALL', error: writeError.message, phase: 'write-phase3'});
  }
  
  SpreadsheetApp.flush();
  
  // STEP 5: Apply formatting to all tickers at once
  Logger.log('Step 5: Applying formatting...');
  try {
    // Apply percentage formatting to columns H, K, V, Z for all data rows
    const numRows = tickers.length;
    calc.getRange(3, 8, numRows, 1).setNumberFormat('0.00%');  // H: Change %
    calc.getRange(3, 11, numRows, 1).setNumberFormat('0.00%'); // K: ATH Diff %
    calc.getRange(3, 22, numRows, 1).setNumberFormat('0.00%'); // V: Stoch %K
    calc.getRange(3, 26, numRows, 1).setNumberFormat('0.00%'); // Z: Bollinger %B
    Logger.log('Percentage formatting applied to all tickers');
  } catch (formatError) {
    Logger.log(`Error applying formatting: ${formatError.message}`);
  }
  
  SpreadsheetApp.flush();
  
  // Apply Bloomberg-style formatting to data rows
  Logger.log('Applying Bloomberg-style formatting to data rows...');
  applyBloombergDataFormatting(calc, tickers.length);
  SpreadsheetApp.flush();
  
  // Calculate total processing time
  const endTime = new Date();
  const elapsedSeconds = ((endTime - startTime) / 1000).toFixed(2);
  
  Logger.log(`Formula generation completed in ${elapsedSeconds}s`);
  
  // Display completion summary
  displaySummary(tickers.length, errors);
}

function generateTickerFormulas(ticker, row, index, BLOCK, SEP, useLongTermSignal) {
  try {
    const t = String(ticker || "").trim().toUpperCase();
    
    if (!t) {
      throw new Error('Empty ticker symbol');
    }
    
    // DATA block start (each ticker is 7 cols in DATA)
    const tDS = (index * BLOCK) + 1;
    const dateCol = columnToLetter(tDS + 0);
    const openCol = columnToLetter(tDS + 1);
    const highCol = columnToLetter(tDS + 2);
    const lowCol = columnToLetter(tDS + 3);
    const closeCol = columnToLetter(tDS + 4);
    const volCol = columnToLetter(tDS + 5);
    
    // Cached fundamentals in DATA row 3
    const athCell = `DATA!${columnToLetter(tDS + 1)}3`;
    const peCell = `DATA!${columnToLetter(tDS + 3)}3`;
    const epsCell = `DATA!${columnToLetter(tDS + 5)}3`;
    
    // Rolling window anchors
    const lastRowCount = `COUNTA(DATA!${closeCol}$5:${closeCol})`;
    const lastAbsRow = `(4+${lastRowCount})`;
    
    // Build all formulas for this ticker (33 formulas: B-AH)
    // CORRECT COLUMN ORDER per headers:
    // B=MARKET RATING, C=DECISION, D=SIGNAL, E=PATTERNS, F=CONSENSUS PRICE, G=Price, H=Change%, I=Vol Trend,
    // J=ATH TRUE, K=ATH Diff%, L=ATH ZONE, M=FUNDAMENTAL,
    // N=Trend State, O=SMA 20, P=SMA 50, Q=SMA 200,
    // R=RSI, S=MACD Hist, T=Divergence, U=ADX, V=Stoch %K,
    // W=VOL REGIME, X=BBP SIGNAL, Y=ATR, Z=Bollinger %B,
    // AA=Target, AB=R:R Quality, AC=Support, AD=Resistance, AE=ATR STOP, AF=ATR TARGET,
    // AG=POSITION SIZE, AH=LAST STATE
    
    const formulas = [
      buildMarketRatingFormula(row, SEP),                                 // B: MARKET RATING (from INPUT D)
      `=DECISION($A${row}${SEP}$D${row}${SEP}$E${row})`,                 // C: DECISION (custom function with dependencies)
      `=SIGNAL($A${row}${SEP}$G${row}${SEP}$I${row}${SEP}$O${row}${SEP}$P${row}${SEP}$Q${row}${SEP}$AC${row}${SEP}$AD${row})`,     // D: SIGNAL (custom function - price triggers dependent indicators)
      `=GETPATTERNS($A${row})`,                                           // E: PATTERNS (old D formula - pattern detection)
      buildConsensusPriceFormula(row, SEP),                               // F: CONSENSUS PRICE (from INPUT E)
      `=ROUND(IFERROR(GOOGLEFINANCE("${t}"${SEP}"price")${SEP}0)${SEP}2)`, // G: Price (old E formula)
      `=IFERROR(GOOGLEFINANCE("${t}"${SEP}"changepct")/100${SEP}0)`,      // H: Change % (shifted from F)
      buildRVOLFormula(row, volCol, lastRowCount, SEP),                   // I: Vol Trend (shifted from G)
      `=IFERROR(${athCell}${SEP}0)`,                                      // J: ATH (TRUE) (shifted from H) - reads from DATA sheet
      `=IFERROR(($G${row}-$J${row})/MAX(0.01${SEP}$J${row})${SEP}0)`,    // K: ATH Diff % (shifted from I) - uses J not K!
      buildATHZoneFormula(row, SEP),                                      // L: ATH ZONE (shifted from J) - uses K not L!
      buildFundamentalFormula(row, peCell, epsCell, SEP),                 // M: FUNDAMENTAL (shifted from K) - uses K not M!
      `=IF($G${row}>$Q${row}${SEP}"BULL"${SEP}"BEAR")`,                   // N: Trend State (shifted from L) - uses Q (SMA 200) not R!
      buildSMAFormula(closeCol, lastRowCount, 20, SEP),                   // O: SMA 20 (shifted from M)
      buildSMAFormula(closeCol, lastRowCount, 50, SEP),                   // P: SMA 50 (shifted from N)
      buildSMAFormula(closeCol, lastRowCount, 200, SEP),                  // Q: SMA 200 (shifted from O)
      `=LIVERSI(DATA!${closeCol}$5:${closeCol}${SEP}$G${row})`,          // R: RSI (shifted from P)
      `=LIVEMACD(DATA!${closeCol}$5:${closeCol}${SEP}$G${row})`,         // S: MACD Hist (shifted from Q)
      buildDivergenceFormula(row, closeCol, lastAbsRow, SEP),             // T: Divergence (shifted from R) - uses S not T!
      `=IFERROR(LIVEADX(DATA!${highCol}$5:${highCol}${SEP}DATA!${lowCol}$5:${lowCol}${SEP}DATA!${closeCol}$5:${closeCol}${SEP}$G${row})${SEP}0)`, // U: ADX (shifted from S)
      `=LIVESTOCHK(DATA!${highCol}$5:${highCol}${SEP}DATA!${lowCol}$5:${lowCol}${SEP}DATA!${closeCol}$5:${closeCol}${SEP}$G${row})`, // V: Stoch %K (shifted from T)
      buildVolRegimeFormula(row, SEP),                                    // W: VOL REGIME (shifted from U)
      buildBBPSignalFormula(row, SEP),                                    // X: BBP SIGNAL (shifted from V) - uses R and Z not S and AA!
      `=IFERROR(LIVEATR(DATA!${highCol}$5:${highCol}${SEP}DATA!${lowCol}$5:${lowCol}${SEP}DATA!${closeCol}$5:${closeCol}${SEP}$G${row})${SEP}0)`, // Y: ATR (shifted from W)
      buildBBPFormula(row, closeCol, lastRowCount, SEP),                  // Z: Bollinger %B (shifted from X) - uses O (SMA 20) not P!
      `=ROUND(MAX($AD${row}${SEP}$G${row}+(($G${row}-$AC${row})*3))${SEP}2)`, // AA: Target (shifted from Y)
      buildRRFormula(row, SEP),                                           // AB: R:R Quality (shifted from Z)
      buildSupportFormula(row, lowCol, lastRowCount, SEP),                // AC: Support (shifted from AA) - uses U (ADX) not W!
      buildResistanceFormula(row, highCol, lastRowCount, SEP),            // AD: Resistance (shifted from AB) - uses U (ADX) not W!
      `=ROUND(MAX($AC${row}${SEP}$G${row}-($Y${row}*2))${SEP}2)`,        // AE: ATR STOP (shifted from AC)
      `=ROUND($G${row}+($Y${row}*3)${SEP}2)`,                             // AF: ATR TARGET (shifted from AD)
      buildPositionSizeFormula(row, SEP),                                 // AG: POSITION SIZE (shifted from AE) - uses K and AB not L and AC!
      `=IF($A${row}=""${SEP}""${SEP}$C${row})`                            // AH: LAST STATE (shifted from AF) (references DECISION)
    ];
    
    // Validate that we have exactly 33 formulas
    if (formulas.length !== 33) {
      throw new Error(`Formula count mismatch: expected 33, got ${formulas.length}`);
    }
    
    // Validate that all formulas are strings
    for (let i = 0; i < formulas.length; i++) {
      if (typeof formulas[i] !== 'string') {
        throw new Error(`Formula at index ${i} is not a string: ${typeof formulas[i]}`);
      }
    }
    
    return formulas;
    
  } catch (error) {
    Logger.log(`generateTickerFormulas error for ${ticker}: ${error.message}`);
    Logger.log(`Stack: ${error.stack}`);
    throw error; // Re-throw to be caught by caller
  }
}

// Helper formula builders
function buildSignalFormula(row, SEP, useLongTermSignal) {
  // CORRECT Column references per new column order:
  // G=Price, I=Vol Trend, J=ATH TRUE, K=ATH Diff%, O=SMA 20, P=SMA 50, Q=SMA 200,
  // R=RSI, S=MACD Hist, U=ADX, V=Stoch %K, Y=ATR, Z=Bollinger %B, AC=Support, AD=Resistance
  
  if (useLongTermSignal) {
    // LONG-TERM INVESTMENT MODE - Conservative, trend-following approach
    return `=IF(OR(ISBLANK($G${row})${SEP}$G${row}=0)${SEP}"LOADING"${SEP}IFS($G${row}<$AC${row}${SEP}"STOP OUT"${SEP}$G${row}<$Q${row}${SEP}"RISK OFF"${SEP}AND($G${row}>$Q${row}${SEP}$P${row}>$Q${row}${SEP}$R${row}>=30${SEP}$R${row}<=40${SEP}$S${row}>0${SEP}$U${row}>=20${SEP}$I${row}>=1.5)${SEP}"STRONG BUY"${SEP}AND($G${row}>$Q${row}${SEP}$P${row}>$Q${row}${SEP}$R${row}>40${SEP}$R${row}<=50${SEP}$S${row}>0${SEP}$U${row}>=15)${SEP}"BUY"${SEP}AND($G${row}>$Q${row}${SEP}$R${row}>=35${SEP}$R${row}<=55${SEP}$G${row}>=$P${row}*0.95${SEP}$G${row}<=$P${row}*1.05)${SEP}"ACCUMULATE"${SEP}AND($R${row}<=30${SEP}$G${row}>$AC${row})${SEP}"OVERSOLD WATCH"${SEP}OR($R${row}>=70${SEP}$Z${row}>=0.85${SEP}$G${row}>=$AD${row}*0.98)${SEP}"TRIM"${SEP}AND($G${row}>$Q${row}${SEP}$R${row}>40${SEP}$R${row}<70)${SEP}"HOLD"${SEP}TRUE${SEP}"NEUTRAL"))`;
  } else {
    // TRADE MODE - Momentum and breakout focused
    return `=IF(OR(ISBLANK($G${row})${SEP}$G${row}=0)${SEP}"LOADING"${SEP}IFS($G${row}<$AC${row}${SEP}"STOP OUT"${SEP}AND($Y${row}>IFERROR(AVERAGE(OFFSET($Y${row}${SEP}-MIN(20${SEP}ROW($Y${row})-1)${SEP}0${SEP}MIN(20${SEP}ROW($Y${row})-1)))${SEP}$Y${row})*1.5${SEP}$I${row}>=2.0${SEP}$G${row}>=$AD${row}*1.01)${SEP}"VOLATILITY BREAKOUT"${SEP}AND($I${row}>=1.5${SEP}$G${row}>=$AD${row}*1.02)${SEP}"BREAKOUT"${SEP}AND($K${row}>=-0.01${SEP}$I${row}>=2.0${SEP}$U${row}>=25)${SEP}"ATH BREAKOUT"${SEP}AND($G${row}>$P${row}${SEP}$S${row}>0${SEP}$U${row}>=20)${SEP}"MOMENTUM"${SEP}AND($V${row}<=20${SEP}$S${row}>0${SEP}$G${row}>$AC${row})${SEP}"OVERSOLD REVERSAL"${SEP}AND($Y${row}<IFERROR(AVERAGE(OFFSET($Y${row}${SEP}-MIN(20${SEP}ROW($Y${row})-1)${SEP}0${SEP}MIN(20${SEP}ROW($Y${row})-1)))${SEP}$Y${row})*0.7${SEP}$U${row}<15${SEP}ABS($Z${row}-0.5)<0.2)${SEP}"VOLATILITY SQUEEZE"${SEP}AND($U${row}<15${SEP}$G${row}>=$AC${row}*0.98${SEP}$G${row}<=$AC${row}*1.02)${SEP}"RANGE SUPPORT BUY"${SEP}OR($R${row}>=70${SEP}$Z${row}>=0.9)${SEP}"OVERBOUGHT"${SEP}$G${row}<$Q${row}${SEP}"RISK OFF"${SEP}AND($U${row}<15${SEP}$G${row}>$AC${row})${SEP}"RANGE"${SEP}TRUE${SEP}"NEUTRAL"))`;
  }
}

function buildFundamentalFormula(row, peCell, epsCell, SEP) {
  // ATH Diff % is in column K (not M!)
  return `=IFERROR(LET(peRaw${SEP}${peCell}${SEP}epsRaw${SEP}${epsCell}${SEP}athDiffRaw${SEP}$K${row}${SEP}pe${SEP}IFERROR(VALUE(REGEXREPLACE(TO_TEXT(peRaw)${SEP}"[^0-9\\.\\-]"${SEP}""))${SEP}"")${SEP}eps${SEP}IFERROR(VALUE(REGEXREPLACE(TO_TEXT(epsRaw)${SEP}"[^0-9\\.\\-]"${SEP}""))${SEP}"")${SEP}athDiff${SEP}IFERROR(VALUE(REGEXREPLACE(TO_TEXT(athDiffRaw)${SEP}"[^0-9\\.\\-]"${SEP}""))/100${SEP}"")${SEP}IFS(OR(pe=""${SEP}eps="")${SEP}"FAIR"${SEP}eps<=0${SEP}"ZOMBIE"${SEP}AND(pe>=60${SEP}athDiff<>""${SEP}athDiff>=-0.08)${SEP}"PRICED FOR PERFECTION"${SEP}pe>=35${SEP}"EXPENSIVE"${SEP}AND(pe>0${SEP}pe<=25${SEP}eps>=0.5)${SEP}"VALUE"${SEP}AND(pe>25${SEP}pe<35${SEP}eps>=0.5)${SEP}"FAIR"${SEP}TRUE${SEP}"FAIR"))${SEP}"FAIR")`;
}

function buildDecisionFormula(row, SEP, useLongTermSignal) {
  // CRITICAL: DECISION (C) uses SIGNAL (D) + PATTERNS (E)
  // FUNDAMENTAL (M) is informational but does NOT drive DECISION logic
  
  const tagExpr = `UPPER(IFERROR(INDEX(INPUT!$C$3:$C${SEP}MATCH($A${row}${SEP}INPUT!$A$3:$A${SEP}0))${SEP}""))`;
  const purchasedExpr = `ISNUMBER(SEARCH("PURCHASED"${SEP}${tagExpr}))`;
  
  // Pattern analysis helpers - Updated to use short forms - NOW REFERENCES COLUMN E (PATTERNS)
  const hasBullishPattern = `OR(ISNUMBER(SEARCH("ASC_TRI"${SEP}$E${row}))${SEP}ISNUMBER(SEARCH("BRKOUT"${SEP}$E${row}))${SEP}ISNUMBER(SEARCH("DBL_BTM"${SEP}$E${row}))${SEP}ISNUMBER(SEARCH("INV_H&S"${SEP}$E${row}))${SEP}ISNUMBER(SEARCH("CUP_HDL"${SEP}$E${row})))`;
  const hasBearishPattern = `OR(ISNUMBER(SEARCH("DESC_TRI"${SEP}$E${row}))${SEP}ISNUMBER(SEARCH("H&S"${SEP}$E${row}))${SEP}ISNUMBER(SEARCH("DBL_TOP"${SEP}$E${row})))`;
  const hasPattern = `NOT(OR($E${row}=""${SEP}$E${row}="—"))`;
  
  if (useLongTermSignal) {
    // Long-term investment mode: SIGNAL (D) + PATTERNS (E) (no FUNDAMENTAL in logic)
    return `=IF($A${row}=""${SEP}""${SEP}IF($D${row}="LOADING"${SEP}"LOADING"${SEP}IF(${purchasedExpr}${SEP}` +
      // For PURCHASED positions
      `IFS(` +
      `OR($D${row}="STOP OUT"${SEP}$D${row}="RISK OFF")${SEP}"🔴 EXIT"${SEP}` +
      `AND($D${row}="TRIM"${SEP}${hasPattern}${SEP}${hasBearishPattern})${SEP}"🟠 TRIM (PATTERN CONFIRMED)"${SEP}` +
      `$D${row}="TRIM"${SEP}"🟠 TRIM"${SEP}` +
      `AND(OR($D${row}="STRONG BUY"${SEP}$D${row}="BUY"${SEP}$D${row}="ACCUMULATE")${SEP}${hasPattern}${SEP}${hasBullishPattern})${SEP}"🟢 ADD (PATTERN CONFIRMED)"${SEP}` +
      `AND(OR($D${row}="STRONG BUY"${SEP}$D${row}="BUY"${SEP}$D${row}="ACCUMULATE")${SEP}${hasPattern}${SEP}${hasBearishPattern})${SEP}"⚠️ HOLD (PATTERN CONFLICT)"${SEP}` +
      `OR($D${row}="STRONG BUY"${SEP}$D${row}="BUY"${SEP}$D${row}="ACCUMULATE")${SEP}"🟢 ADD"${SEP}` +
      `$D${row}="HOLD"${SEP}"⚖️ HOLD"${SEP}` +
      `TRUE${SEP}"⚖️ HOLD"` +
      `)${SEP}` +
      // For NON-PURCHASED positions
      `IFS(` +
      `OR($D${row}="STOP OUT"${SEP}$D${row}="RISK OFF")${SEP}"🔴 AVOID"${SEP}` +
      `AND($D${row}="STRONG BUY"${SEP}${hasPattern}${SEP}${hasBullishPattern})${SEP}"🟢 STRONG BUY (PATTERN CONFIRMED)"${SEP}` +
      `AND(OR($D${row}="STRONG BUY"${SEP}$D${row}="BUY")${SEP}${hasPattern}${SEP}${hasBearishPattern})${SEP}"⚠️ CAUTION (PATTERN CONFLICT)"${SEP}` +
      `$D${row}="STRONG BUY"${SEP}"🟢 STRONG BUY"${SEP}` +
      `$D${row}="BUY"${SEP}"🟢 BUY"${SEP}` +
      `$D${row}="ACCUMULATE"${SEP}"🟢 ACCUMULATE"${SEP}` +
      `$D${row}="OVERSOLD WATCH"${SEP}"🟡 WATCH (OVERSOLD)"${SEP}` +
      `$D${row}="TRIM"${SEP}"⏳ WAIT (EXTENDED)"${SEP}` +
      `$D${row}="HOLD"${SEP}"⚖️ WATCH"${SEP}` +
      `TRUE${SEP}"⚪ NEUTRAL"` +
      `)` +
      `)))`;
  } else {
    // Trade mode: SIGNAL (D) + PATTERNS (E) (no FUNDAMENTAL in logic)
    return `=IF($A${row}=""${SEP}""${SEP}LET(` +
      `tag${SEP}UPPER(IFERROR(INDEX(INPUT!$C$3:$C${SEP}MATCH($A${row}${SEP}INPUT!$A$3:$A${SEP}0))${SEP}""))${SEP}` +
      `purchased${SEP}REGEXMATCH(tag${SEP}"(^|,|\\\\s)PURCHASED(\\\\s|,|$)")${SEP}` +
      `bullishPat${SEP}${hasBullishPattern}${SEP}` +
      `bearishPat${SEP}${hasBearishPattern}${SEP}` +
      `hasPat${SEP}${hasPattern}${SEP}` +
      `IFS(` +
      // Stop-out check - Price below Support
      `AND(IFERROR(VALUE($G${row})${SEP}0)>0${SEP}IFERROR(VALUE($AC${row})${SEP}0)>0${SEP}IFERROR(VALUE($G${row})${SEP}0)<IFERROR(VALUE($AC${row})${SEP}0))${SEP}"🔴 STOP OUT"${SEP}` +
      // Pattern-confirmed strong signals
      `AND(NOT(purchased)${SEP}$D${row}="VOLATILITY BREAKOUT"${SEP}hasPat${SEP}bullishPat)${SEP}"🟢 STRONG TRADE LONG (PATTERN CONFIRMED)"${SEP}` +
      `AND(NOT(purchased)${SEP}OR($D${row}="BREAKOUT"${SEP}$D${row}="ATH BREAKOUT")${SEP}hasPat${SEP}bullishPat)${SEP}"🟢 TRADE LONG (PATTERN CONFIRMED)"${SEP}` +
      // Pattern conflicts
      `AND(NOT(purchased)${SEP}OR($D${row}="VOLATILITY BREAKOUT"${SEP}$D${row}="BREAKOUT"${SEP}$D${row}="ATH BREAKOUT"${SEP}$D${row}="MOMENTUM")${SEP}hasPat${SEP}bearishPat)${SEP}"⚠️ CAUTION (PATTERN CONFLICT)"${SEP}` +
      // Standard signals without pattern consideration
      `AND(NOT(purchased)${SEP}$D${row}="VOLATILITY BREAKOUT")${SEP}"🟢 STRONG TRADE LONG"${SEP}` +
      `AND(NOT(purchased)${SEP}OR($D${row}="BREAKOUT"${SEP}$D${row}="ATH BREAKOUT"))${SEP}"🟢 TRADE LONG"${SEP}` +
      `AND(NOT(purchased)${SEP}$D${row}="MOMENTUM")${SEP}"🟡 ACCUMULATE"${SEP}` +
      `AND(NOT(purchased)${SEP}$D${row}="OVERSOLD REVERSAL")${SEP}"🟢 BUY DIP"${SEP}` +
      `AND(NOT(purchased)${SEP}$D${row}="RANGE SUPPORT BUY")${SEP}"🟡 RANGE BUY"${SEP}` +
      `AND(NOT(purchased)${SEP}$D${row}="VOLATILITY SQUEEZE")${SEP}"⏳ WAIT FOR BREAKOUT"${SEP}` +
      // Purchased position management
      `AND(purchased${SEP}OR($D${row}="OVERBOUGHT"${SEP}IFERROR(VALUE($G${row})${SEP}0)>=IFERROR(VALUE($AD${row})${SEP}0)*0.98))${SEP}"🟠 TAKE PROFIT"${SEP}` +
      `AND(purchased${SEP}$D${row}="RISK OFF")${SEP}"🔴 RISK OFF"${SEP}` +
      `AND(NOT(purchased)${SEP}$D${row}="RISK OFF")${SEP}"🔴 AVOID"${SEP}` +
      // Default holds
      `purchased${SEP}"⚖️ HOLD"${SEP}` +
      `TRUE${SEP}"⚪ NEUTRAL"` +
      `)))`;
  }
}

function buildRVOLFormula(row, volCol, lastRowCount, SEP) {
  return `=ROUND(IFERROR(OFFSET(DATA!${volCol}$5${SEP}${lastRowCount}-1${SEP}0)/AVERAGE(OFFSET(DATA!${volCol}$5${SEP}${lastRowCount}-20${SEP}0${SEP}20))${SEP}1)${SEP}2)`;
}

function buildRRFormula(row, SEP) {
  // G=Price, Y=ATR, AC=Support, AD=Resistance
  return `=IF(OR($G${row}<=$AC${row}${SEP}$G${row}=0)${SEP}0${SEP}ROUND(MAX(0${SEP}$AD${row}-$G${row})/MAX($Y${row}*0.5${SEP}$G${row}-$AC${row})${SEP}2))`;
}

function buildSMAFormula(closeCol, lastRowCount, period, SEP) {
  return `=ROUND(IFERROR(AVERAGE(OFFSET(DATA!${closeCol}$5${SEP}${lastRowCount}-${period}${SEP}0${SEP}${period}))${SEP}0)${SEP}2)`;
}

function buildDivergenceFormula(row, closeCol, lastAbsRow, SEP) {
  // G=Price, R=RSI
  return `=IFERROR(IFS(AND($G${row}<INDEX(DATA!${closeCol}:${closeCol}${SEP}${lastAbsRow}-14)${SEP}$R${row}>50)${SEP}"BULL DIV"${SEP}AND($G${row}>INDEX(DATA!${closeCol}:${closeCol}${SEP}${lastAbsRow}-14)${SEP}$R${row}<50)${SEP}"BEAR DIV"${SEP}TRUE${SEP}"—")${SEP}"—")`;
}

function buildSupportFormula(row, lowCol, lastRowCount, SEP) {
  // U=ADX
  return `=ROUND(IFERROR(LET(win${SEP}IFS($U${row}<20${SEP}10${SEP}$U${row}<35${SEP}22${SEP}TRUE${SEP}40)${SEP}n${SEP}${lastRowCount}${SEP}start${SEP}MAX(0${SEP}n-win)${SEP}len${SEP}MIN(win${SEP}n)${SEP}rng${SEP}IF(len<=0${SEP}OFFSET(DATA!${lowCol}$5${SEP}0${SEP}0)${SEP}OFFSET(DATA!${lowCol}$5${SEP}start${SEP}0${SEP}len))${SEP}out${SEP}IF(COUNTA(rng)<3${SEP}IFERROR(MIN(rng)${SEP}0)${SEP}PERCENTILE.INC(rng${SEP}0.15))${SEP}out)${SEP}0)${SEP}2)`;
}

function buildResistanceFormula(row, highCol, lastRowCount, SEP) {
  // U=ADX
  return `=ROUND(IFERROR(LET(win${SEP}IFS($U${row}<20${SEP}10${SEP}$U${row}<35${SEP}22${SEP}TRUE${SEP}40)${SEP}n${SEP}${lastRowCount}${SEP}start${SEP}MAX(0${SEP}n-win)${SEP}len${SEP}MIN(win${SEP}n)${SEP}rng${SEP}IF(len<=0${SEP}OFFSET(DATA!${highCol}$5${SEP}0${SEP}0)${SEP}OFFSET(DATA!${highCol}$5${SEP}start${SEP}0${SEP}len))${SEP}out${SEP}IF(COUNTA(rng)<3${SEP}IFERROR(MAX(rng)${SEP}0)${SEP}PERCENTILE.INC(rng${SEP}0.85))${SEP}out)${SEP}0)${SEP}2)`;
}

function buildBBPFormula(row, closeCol, lastRowCount, SEP) {
  // G=Price, O=SMA 20
  return `=ROUND(IFERROR((($G${row}-$O${row})/(4*STDEV(OFFSET(DATA!${closeCol}$5${SEP}${lastRowCount}-20${SEP}0${SEP}20))))+0.5${SEP}0.5)${SEP}2)`;
}

function buildPositionSizeFormula(row, SEP) {
  // G=Price, K=ATH Diff%, Y=ATR, AB=R:R Quality
  return `=IF($A${row}=""${SEP}""${SEP}LET(riskReward${SEP}$AB${row}${SEP}atrRisk${SEP}$Y${row}/$G${row}${SEP}athRisk${SEP}IF($K${row}>=-0.05${SEP}0.8${SEP}1.0)${SEP}volRegimeRisk${SEP}IFS(atrRisk<=0.02${SEP}1.2${SEP}atrRisk<=0.05${SEP}1.0${SEP}atrRisk<=0.08${SEP}0.7${SEP}TRUE${SEP}0.5)${SEP}baseSize${SEP}0.02${SEP}rrMultiplier${SEP}IF(riskReward>=3${SEP}1.5${SEP}IF(riskReward>=2${SEP}1.0${SEP}0.5))${SEP}finalSize${SEP}MIN(0.08${SEP}baseSize*rrMultiplier*volRegimeRisk*athRisk)${SEP}TEXT(finalSize${SEP}"0.0%")&" (Vol: "&IFS(atrRisk<=0.02${SEP}"LOW"${SEP}atrRisk<=0.05${SEP}"NORM"${SEP}atrRisk<=0.08${SEP}"HIGH"${SEP}TRUE${SEP}"EXTR")&")"))`;
}

function buildVolRegimeFormula(row, SEP) {
  // Y=ATR, G=Price
  return `=IFS($Y${row}/$G${row}<=0.02${SEP}"LOW VOL"${SEP}$Y${row}/$G${row}<=0.05${SEP}"NORMAL VOL"${SEP}$Y${row}/$G${row}<=0.08${SEP}"HIGH VOL"${SEP}TRUE${SEP}"EXTREME VOL")`;
}

function buildATHZoneFormula(row, SEP) {
  // K=ATH Diff %
  return `=IFS($K${row}>=-0.02${SEP}"AT ATH"${SEP}$K${row}>=-0.05${SEP}"NEAR ATH"${SEP}$K${row}>=-0.15${SEP}"RESISTANCE ZONE"${SEP}$K${row}>=-0.30${SEP}"PULLBACK ZONE"${SEP}$K${row}>=-0.50${SEP}"CORRECTION ZONE"${SEP}TRUE${SEP}"DEEP VALUE ZONE")`;
}

function buildBBPSignalFormula(row, SEP) {
  // R=RSI, Z=Bollinger %B, G=Price, Q=SMA 200, AC=Support
  return `=IFS(AND($Z${row}>=0.9${SEP}$R${row}>=70)${SEP}"EXTREME OVERBOUGHT"${SEP}AND($Z${row}<=0.1${SEP}$R${row}<=30)${SEP}"EXTREME OVERSOLD"${SEP}AND($Z${row}>=0.8${SEP}$G${row}>$Q${row})${SEP}"MOMENTUM STRONG"${SEP}AND($Z${row}<=0.2${SEP}$G${row}>$AC${row})${SEP}"MEAN REVERSION"${SEP}TRUE${SEP}"NEUTRAL")`;
}

function buildMarketRatingFormula(row, SEP) {
  // Reference MARKET RATING from INPUT sheet column D
  return `=IFERROR(INDEX(INPUT!$D$3:$D${SEP}MATCH($A${row}${SEP}INPUT!$A$3:$A${SEP}0))${SEP}"—")`;
}

function buildConsensusPriceFormula(row, SEP) {
  // Reference CONSENSUS PRICE from INPUT sheet column E
  return `=IFERROR(INDEX(INPUT!$E$3:$E${SEP}MATCH($A${row}${SEP}INPUT!$A$3:$A${SEP}0))${SEP}0)`;
}

/**
 * Process a single ticker: detect patterns, generate formulas, write to sheet
 * Implements comprehensive error handling for each phase of processing
 * @param {string} ticker - Ticker symbol
 * @param {number} row - Row number in CALCULATIONS sheet
 * @param {number} index - Index in tickers array
 * @param {number} BLOCK - DATA block width
 * @param {string} SEP - Separator character
 * @param {boolean} useLongTermSignal - Whether to use long-term signals
 * @param {Sheet} dataSheet - DATA sheet reference
 * @param {Sheet} calc - CALCULATIONS sheet reference
 * @returns {Object} {success: boolean, error: string|null, phase: string|null}
 */
function processTickerFormulas(ticker, row, index, BLOCK, SEP, useLongTermSignal, dataSheet, calc) {
  try {
    // PHASE 1: Pattern Detection
    let patternString = '';
    try {
      const priceData = getPriceDataForTicker(dataSheet, ticker, index, BLOCK);
      
      if (!priceData || priceData.length === 0) {
        Logger.log(`${ticker}: No price data available`);
        setCachedPattern(ticker, '');
      } else {
        const patterns = detectPatterns(priceData, {minBars: 100, minConfidence: 60});
        patternString = formatPatternsForSheet(patterns);
        
        // Cache the pattern result for use by GETPATTERNS formula
        setCachedPattern(ticker, patternString);
        
        Logger.log(`${ticker}: Cached ${patterns.length} patterns - ${patternString || 'none'}`);
      }
    } catch (patternError) {
      Logger.log(`${ticker}: Error detecting patterns - ${patternError.message}`);
      setCachedPattern(ticker, '');
      // Continue processing - pattern detection failure is not fatal
      // Formulas will still be written, pattern column will be empty
    }
    
    // PHASE 2: Formula Generation
    let formulas;
    try {
      formulas = generateTickerFormulas(ticker, row, index, BLOCK, SEP, useLongTermSignal);
      
      // Validate that formulas array has correct length
      if (!formulas || !Array.isArray(formulas) || formulas.length !== 31) {
        throw new Error(`Invalid formulas array: expected 31 elements, got ${formulas ? formulas.length : 'null'}`);
      }
      
      Logger.log(`${ticker}: Formulas generated successfully`);
    } catch (formulaError) {
      Logger.log(`${ticker}: Error generating formulas - ${formulaError.message}`);
      return {
        success: false,
        error: formulaError.message,
        phase: 'formula'
      };
    }
    
    // PHASE 3: Formula Writing (Staggered)
    try {
      // Phase 1: Write formulas for columns E-AF (indices 3-30 in formulas array)
      // These are the main calculation formulas that don't depend on B, C, D
      const phase1Formulas = formulas.slice(3);
      calc.getRange(row, 5, 1, 28).setFormulas([phase1Formulas]);
      Logger.log(`${ticker}: Phase 1 complete - wrote formulas for columns E-AF (28 columns)`);
    } catch (writeError) {
      Logger.log(`${ticker}: Error writing Phase 1 formulas - ${writeError.message}`);
      return {
        success: false,
        error: writeError.message,
        phase: 'write-phase1'
      };
    }
    
    // Delay after Phase 1: Allow calculation engine to process main formulas
    try {
      Logger.log(`${ticker}: Applying ${DELAY_AFTER_MAIN_FORMULAS}ms delay after Phase 1...`);
      Utilities.sleep(DELAY_AFTER_MAIN_FORMULAS);
    } catch (delayError) {
      Logger.log(`${ticker}: Error during Phase 1 delay - ${delayError.message}`);
      // Delay error is not fatal - continue with next phase
    }
    
    // Phase 2: Write formulas for columns C-D (indices 1-2 in formulas array)
    // These formulas depend on price data and pattern detection
    try {
      // Extract formulas for indices 1-2 from formulas array
      const phase2Formulas = formulas.slice(1, 3);
      
      // Write to range starting at column 3 (C) with 2 columns
      calc.getRange(row, 3, 1, 2).setFormulas([phase2Formulas]);
      
      // Add logging for Phase 2 completion
      Logger.log(`${ticker}: Phase 2 complete - wrote formulas for columns C-D (2 columns)`);
    } catch (writeError) {
      Logger.log(`${ticker}: Error writing Phase 2 formulas - ${writeError.message}`);
      return {
        success: false,
        error: writeError.message,
        phase: 'write-phase2'
      };
    }
    
    // Delay after Phase 2: Allow calculation engine to process C-D formulas
    try {
      Logger.log(`${ticker}: Applying ${DELAY_AFTER_CD_FORMULAS}ms delay after Phase 2...`);
      Utilities.sleep(DELAY_AFTER_CD_FORMULAS);
    } catch (delayError) {
      Logger.log(`${ticker}: Error during Phase 2 delay - ${delayError.message}`);
      // Delay error is not fatal - continue with next phase
    }
    
    // Phase 3: Write formula for column B (index 0 in formulas array)
    // This formula depends on many other columns and should be written last
    try {
      const phase3Formula = formulas[0];
      calc.getRange(row, 2, 1, 1).setFormulas([[phase3Formula]]);
      Logger.log(`${ticker}: Phase 3 complete - wrote formula for column B (SIGNAL)`);
    } catch (writeError) {
      Logger.log(`${ticker}: Error writing Phase 3 formula - ${writeError.message}`);
      return {
        success: false,
        error: writeError.message,
        phase: 'write-phase3'
      };
    }
    
    // PHASE 4: Formatting
    try {
      // Apply percentage formatting to columns F, J, U, X
      calc.getRange(row, 6, 1, 1).setNumberFormat('0.00%');  // F: Change %
      calc.getRange(row, 10, 1, 1).setNumberFormat('0.00%'); // J: ATH Diff %
      calc.getRange(row, 21, 1, 1).setNumberFormat('0.00%'); // U: Stoch %K
      calc.getRange(row, 24, 1, 1).setNumberFormat('0.00%'); // X: Bollinger %B
      
      Logger.log(`${ticker}: Formatting applied successfully`);
    } catch (formatError) {
      Logger.log(`${ticker}: Error applying formatting - ${formatError.message}`);
      // Formatting error is not fatal - formulas are already written
      // Continue and mark as successful since formulas are in place
    }
    
    // All phases completed successfully
    return {
      success: true,
      error: null,
      phase: null
    };
    
  } catch (unexpectedError) {
    // Catch any unexpected errors not handled by phase-specific try-catch blocks
    Logger.log(`${ticker}: Unexpected error - ${unexpectedError.message}`);
    Logger.log(`${ticker}: Stack trace - ${unexpectedError.stack}`);
    return {
      success: false,
      error: unexpectedError.message,
      phase: 'unexpected'
    };
  }
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

/**
 * Display progress for ticker processing
 * Shows current progress with ticker name and milestone notifications
 * @param {number} current - Current ticker index (0-based)
 * @param {number} total - Total number of tickers
 * @param {string} ticker - Current ticker symbol
 */
function displayProgress(current, total, ticker) {
  const percentage = Math.round((current / total) * 100);
  const tickerNum = current + 1; // Convert to 1-based for display
  
  // Log progress for each ticker
  Logger.log(`Processing ${tickerNum}/${total}: ${ticker} (${percentage}% complete)`);
  
  // Display milestone notifications at 25%, 50%, 75%
  const milestones = [25, 50, 75];
  for (const milestone of milestones) {
    // Check if we just hit this milestone (within 1% tolerance to avoid missing it)
    const prevPercentage = Math.round(((current - 1) / total) * 100);
    if (percentage >= milestone && prevPercentage < milestone) {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      ss.toast(`${milestone}% complete (${tickerNum}/${total} tickers processed)`, '⏳ Progress', 3);
      break; // Only show one milestone per ticker
    }
  }
}

/**
 * Display processing summary with success/error counts
 * Shows completion notification with detailed error information
 * @param {number} total - Total tickers processed
 * @param {Array<Object>} errors - Array of error objects {ticker, error, phase}
 */
function displaySummary(total, errors) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const successCount = total - errors.length;
  const errorCount = errors.length;
  
  // Build summary message
  let summaryMsg = `✓ Processing complete: ${successCount}/${total} successful`;
  
  if (errorCount > 0) {
    summaryMsg += `, ${errorCount} failed`;
    
    // Log detailed error information
    Logger.log(`\n=== PROCESSING SUMMARY ===`);
    Logger.log(`Successful: ${successCount}`);
    Logger.log(`Failed: ${errorCount}`);
    Logger.log(`\n=== ERROR DETAILS ===`);
    
    errors.forEach((err, index) => {
      Logger.log(`${index + 1}. ${err.ticker} (${err.phase}): ${err.error}`);
    });
    
    // Show toast with error summary
    ss.toast(summaryMsg, '⚠️ Completed with Errors', 5);
  } else {
    // All successful
    Logger.log(`\n=== PROCESSING SUMMARY ===`);
    Logger.log(`All ${total} tickers processed successfully`);
    ss.toast(summaryMsg, '✓ Success', 3);
  }
}

/**
 * OPTIMIZED: Update only SIGNAL (D) and DECISION (C) formulas when DASHBOARD!H1 is toggled
 * This is much faster than regenerating the entire CALCULATIONS sheet
 * Only updates the formulas that depend on the useLongTermSignal flag
 */
function updateSignalFormulas() {
  const startTime = new Date();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  try {
    // Get required sheets
    const calc = ss.getSheetByName("CALCULATIONS");
    const dashboard = ss.getSheetByName("DASHBOARD");
    
    if (!calc || !dashboard) {
      ss.toast('Required sheets (CALCULATIONS or DASHBOARD) not found.', '❌ Error', 3);
      return;
    }

    // Get locale separator
    const locale = (ss.getSpreadsheetLocale() || "").toLowerCase();
    const SEP = (/^(en|en_)/.test(locale)) ? "," : ";";

    // Check if long-term signal mode is enabled
    const useLongTermSignal = dashboard.getRange('H1').getValue() === true;
    
    // Get all tickers from column A (starting from row 3)
    const tickerRange = calc.getRange("A3:A").getValues();
    const tickers = tickerRange.filter(row => row[0] && String(row[0]).trim() !== "").map(row => String(row[0]).trim());
    
    if (tickers.length === 0) {
      ss.toast('No tickers found in CALCULATIONS sheet.', '⚠️ Warning', 3);
      return;
    }

    Logger.log(`Updating signal formulas for ${tickers.length} tickers (mode: ${useLongTermSignal ? 'LONG-TERM' : 'TRADE'})`);
    
    // Build formula arrays for batch update
    const signalFormulas = [];
    const decisionFormulas = [];
    
    for (let i = 0; i < tickers.length; i++) {
      const row = i + 3; // Data starts at row 3
      
      // Generate SIGNAL formula (column D)
      const signalFormula = buildSignalFormula(row, SEP, useLongTermSignal);
      signalFormulas.push([signalFormula]);
      
      // Generate DECISION formula (column C)
      const decisionFormula = buildDecisionFormula(row, SEP, useLongTermSignal);
      decisionFormulas.push([decisionFormula]);
    }
    
    // Update SIGNAL formulas (column D) - all at once
    Logger.log('Writing SIGNAL formulas (column D)...');
    calc.getRange(3, 4, tickers.length, 1).setFormulas(signalFormulas);
    SpreadsheetApp.flush();
    
    // Small delay to allow SIGNAL formulas to calculate
    Utilities.sleep(1000);
    
    // Update DECISION formulas (column C) - all at once
    Logger.log('Writing DECISION formulas (column C)...');
    calc.getRange(3, 3, tickers.length, 1).setFormulas(decisionFormulas);
    SpreadsheetApp.flush();
    
    const elapsed = ((new Date() - startTime) / 1000).toFixed(2);
    Logger.log(`Signal formulas updated in ${elapsed}s`);
    
  } catch (error) {
    Logger.log(`Error in updateSignalFormulas: ${error.stack}`);
    ss.toast(`Failed to update signal formulas: ${error.message}`, '❌ Error', 5);
  }
}

// Export functions for testing (Node.js environment)
if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    generateCalculationsSheet,
    setupHeaders,
    applyBloombergDataFormatting,
    writeTickers,
    writeFormulas,
    generateTickerFormulas,
    processTickerFormulas,
    displayProgress,
    displaySummary,
    columnToLetter,
    updateSignalFormulas,
    // Helper formula builders
    buildSignalFormula,
    buildFundamentalFormula,
    buildDecisionFormula,
    buildRVOLFormula,
    buildRRFormula,
    buildSMAFormula,
    buildDivergenceFormula,
    buildSupportFormula,
    buildResistanceFormula,
    buildBBPFormula,
    buildPositionSizeFormula,
    buildVolRegimeFormula,
    buildATHZoneFormula,
    buildBBPSignalFormula,
    buildMarketRatingFormula,
    buildConsensusPriceFormula
  };
}


// ==============================================================================
// SIGNAL AND DECISION CUSTOM FUNCTIONS
// ==============================================================================

/**
 * Custom Apps Script Functions for SIGNAL and DECISION
 * 
 * These functions replace the complex IFS formulas in CALCULATIONS sheet
 * with cleaner, more maintainable Apps Script code.
 * 
 * Usage in CALCULATIONS sheet:
 * - Column D (SIGNAL): =SIGNAL(A3, G3, I3, O3, P3, Q3, AC3, AD3)
 *   Parameters: ticker, price, volTrend, sma20, sma50, sma200, support, resistance
 *   Note: Price change triggers recalculation of K, R, S, U, V, Y, Z automatically
 * - Column C (DECISION): =DECISION(A3, D3, E3)
 *   Parameters: ticker, signal, patterns (for auto-recalculation)
 * 
 * Benefits:
 * - Single source of truth for logic
 * - No mismatch between CALCULATIONS and DASH_REPORT
 * - Better performance than nested IFS formulas
 * - Easier to maintain and debug
 * - Auto-recalculates when dependent cells change
 * 
 * Technical Analysis Approach:
 * - Uses ATH (All-Time High) instead of 52-week high for breakout signals
 * - ATH is superior: captures true breakout potential without arbitrary time limits
 * - Follows methodology of Mark Minervini (SEPA) and Dan Zanger
 * - Added 52-week low for deep value/contrarian signals (long-term mode)
 * 
 * Version: 2.1 (Auto-recalculation support)
 */

/**
 * SIGNAL - Calculates trading signal for a ticker
 * 
 * @param {string} ticker - Ticker symbol (e.g., "AAPL")
 * @param {number} price - Price (G) - forces recalculation (triggers K, R, S, U, V, Y, Z)
 * @param {number} volTrend - Vol Trend (I) - independent, forces recalculation
 * @param {number} sma20 - SMA 20 (O) - independent, forces recalculation
 * @param {number} sma50 - SMA 50 (P) - independent, forces recalculation
 * @param {number} sma200 - SMA 200 (Q) - independent, forces recalculation
 * @param {number} support - Support (AC) - independent, forces recalculation
 * @param {number} resistance - Resistance (AD) - independent, forces recalculation
 * @return {string} Signal value (e.g., "STRONG BUY", "MOMENTUM", "RISK OFF")
 * @customfunction
 */
function SIGNAL(ticker, price, volTrend, sma20, sma50, sma200, support, resistance) {
  try {
    // Validate input
    if (!ticker || ticker === "") {
      return "—";
    }
    
    const tickerUpper = String(ticker).toUpperCase().trim();
    
    // Get spreadsheet and sheets
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const calc = ss.getSheetByName('CALCULATIONS');
    const data = ss.getSheetByName('DATA');
    const dashboard = ss.getSheetByName('DASHBOARD');
    
    if (!calc || !data) {
      return "ERROR: Required sheets not found";
    }
    
    // Get mode from DASHBOARD H1
    const useLongTermSignal = dashboard ? dashboard.getRange('H1').getValue() === true : false;
    
    // Find ticker row in CALCULATIONS
    const calcData = calc.getDataRange().getValues();
    let tickerRow = -1;
    for (let i = 2; i < calcData.length; i++) { // Start from row 3 (index 2)
      if (String(calcData[i][0]).toUpperCase().trim() === tickerUpper) {
        tickerRow = i;
        break;
      }
    }
    
    if (tickerRow === -1) {
      return "ERROR: Ticker not found";
    }
    
    // Build indicator data object from CALCULATIONS row
    const row = calcData[tickerRow];
    const indicators = {
      price: Number(row[6]) || 0,           // G: Price
      volTrend: Number(row[8]) || 0,        // I: Vol Trend
      athDiff: Number(row[10]) || 0,        // K: ATH Diff%
      sma20: Number(row[14]) || 0,          // O: SMA 20
      sma50: Number(row[15]) || 0,          // P: SMA 50
      sma200: Number(row[16]) || 0,         // Q: SMA 200
      rsi: Number(row[17]) || 0,            // R: RSI
      macdHist: Number(row[18]) || 0,       // S: MACD Hist
      adx: Number(row[20]) || 0,            // U: ADX
      stochK: Number(row[21]) || 0,         // V: Stoch %K
      atr: Number(row[24]) || 0,            // Y: ATR
      bollingerPctB: Number(row[25]) || 0,  // Z: Bollinger %B
      support: Number(row[28]) || 0,        // AC: Support
      resistance: Number(row[29]) || 0,     // AD: Resistance
      fundamental: String(row[12]).trim()   // M: FUNDAMENTAL (for DEEP VALUE)
    };
    
    // Check if data is loaded
    if (indicators.price === 0) {
      return "LOADING";
    }
    
    // Calculate ATR average for VOLATILITY BREAKOUT and VOLATILITY SQUEEZE
    // Find ticker column in DATA sheet (each ticker is 7 columns wide)
    const tickerIndex = tickerRow - 2; // Convert to 0-based index
    const atrAverage = calculateATRAverage_Signal(data, tickerIndex);
    indicators.atrAverage = atrAverage;
    
    // Evaluate signal based on mode
    if (useLongTermSignal) {
      return evaluateSignalLongTerm(indicators);
    } else {
      return evaluateSignalTrade(indicators, atrAverage);
    }
    
  } catch (error) {
    Logger.log("SIGNAL function error: " + error.message);
    return "ERROR";
  }
}

/**
 * Calculate 20-period average ATR from DATA sheet
 * @param {Sheet} dataSheet - DATA sheet reference
 * @param {number} tickerIndex - Ticker index (0-based)
 * @return {number} Average ATR over last 20 periods
 * @private
 */
function calculateATRAverage_Signal(dataSheet, tickerIndex) {
  try {
    const BLOCK = 7; // Each ticker occupies 7 columns in DATA
    const atrColIndex = (tickerIndex * BLOCK) + 6; // ATR is 7th column in each block (0-based: index 6)
    
    // Get ATR column data (starting from row 5, which is where data starts)
    const atrColumn = dataSheet.getRange(5, atrColIndex + 1, dataSheet.getLastRow() - 4, 1).getValues();
    
    // Filter out empty/zero values and get last 20 valid ATR values
    const validATRs = [];
    for (let i = atrColumn.length - 1; i >= 0 && validATRs.length < 20; i--) {
      const atrValue = Number(atrColumn[i][0]);
      if (atrValue > 0) {
        validATRs.push(atrValue);
      }
    }
    
    // Calculate average
    if (validATRs.length === 0) {
      return 0;
    }
    
    const sum = validATRs.reduce((acc, val) => acc + val, 0);
    return sum / validATRs.length;
    
  } catch (error) {
    Logger.log("calculateATRAverage_Signal error: " + error.message);
    return 0;
  }
}

/**
 * DECISION - Calculates trading decision for a ticker
 * 
 * @param {string} ticker - Ticker symbol (e.g., "AAPL")
 * @param {string} signal - SIGNAL value (from column D) - forces recalculation
 * @param {string} patterns - PATTERNS value (from column E) - forces recalculation
 * @return {string} Decision value (e.g., "🟢 STRONG BUY", "⚖️ HOLD", "🔴 EXIT")
 * @customfunction
 */
function DECISION(ticker, signal, patterns) {
  try {
    // Validate input
    if (!ticker || ticker === "") {
      return "—";
    }
    
    const tickerUpper = String(ticker).toUpperCase().trim();
    
    // Get spreadsheet and sheets
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const calc = ss.getSheetByName('CALCULATIONS');
    const input = ss.getSheetByName('INPUT');
    const dashboard = ss.getSheetByName('DASHBOARD');
    
    if (!calc || !input) {
      return "ERROR: Required sheets not found";
    }
    
    // Get mode from DASHBOARD H1
    const useLongTermSignal = dashboard ? dashboard.getRange('H1').getValue() === true : false;
    
    // Find ticker row in CALCULATIONS
    const calcData = calc.getDataRange().getValues();
    let tickerRow = -1;
    for (let i = 2; i < calcData.length; i++) { // Start from row 3 (index 2)
      if (String(calcData[i][0]).toUpperCase().trim() === tickerUpper) {
        tickerRow = i;
        break;
      }
    }
    
    if (tickerRow === -1) {
      return "ERROR: Ticker not found";
    }
    
    // Get SIGNAL and PATTERNS from CALCULATIONS
    const row = calcData[tickerRow];
    const signalValue = String(row[3]).trim();      // D: SIGNAL
    const patternsValue = String(row[4]).trim();    // E: PATTERNS
    const price = Number(row[6]) || 0;              // G: Price
    const support = Number(row[28]) || 0;           // AC: Support
    const resistance = Number(row[29]) || 0;        // AD: Resistance
    
    // Check if SIGNAL is loaded
    if (signalValue === "LOADING" || signalValue === "") {
      return "LOADING";
    }
    
    // Check PURCHASED tag from INPUT sheet
    const inputData = input.getDataRange().getValues();
    let isPurchased = false;
    for (let i = 2; i < inputData.length; i++) {
      if (String(inputData[i][0]).toUpperCase().trim() === tickerUpper) {
        const tag = String(inputData[i][2]).toUpperCase().trim();
        isPurchased = (tag === "PURCHASED" || tag === "YES" || tag === "TRUE");
        break;
      }
    }
    
    // Analyze patterns
    const hasPattern = patternsValue !== "" && patternsValue !== "—" && patternsValue !== "-";
    let patternType = "none";
    if (hasPattern) {
      const upperPatterns = patternsValue.toUpperCase();
      const bullishPatterns = ["ASC_TRI", "BRKOUT", "DBL_BTM", "INV_H&S", "CUP_HDL"];
      const bearishPatterns = ["DESC_TRI", "H&S", "DBL_TOP"];
      
      for (const pattern of bullishPatterns) {
        if (upperPatterns.includes(pattern)) {
          patternType = "bullish";
          break;
        }
      }
      
      if (patternType === "none") {
        for (const pattern of bearishPatterns) {
          if (upperPatterns.includes(pattern)) {
            patternType = "bearish";
            break;
          }
        }
      }
    }
    
    // Evaluate decision based on mode
    if (useLongTermSignal) {
      return evaluateDecisionLongTerm(signalValue, isPurchased, hasPattern, patternType);
    } else {
      return evaluateDecisionTrade(signalValue, isPurchased, hasPattern, patternType, price, support, resistance);
    }
    
  } catch (error) {
    Logger.log("DECISION function error: " + error.message);
    return "ERROR";
  }
}

/**
 * Evaluates SIGNAL for LONG-TERM INVESTMENT mode
 * @private
 */
function evaluateSignalLongTerm(ind) {
  // Branch 1: STOP OUT (with 3% buffer to avoid whipsaws)
  if (ind.price < ind.support * 0.97) {
    return "STOP OUT";
  }
  
  // Branch 2: RISK OFF
  if (ind.price < ind.sma200) {
    return "RISK OFF";
  }
  
  // Branch 3: STRONG BUY
  if (ind.price > ind.sma200 && 
      ind.sma50 > ind.sma200 && 
      ind.rsi >= 30 && ind.rsi <= 50 && 
      ind.macdHist > 0 && 
      ind.adx >= 20 && 
      ind.volTrend >= 1.5) {
    return "STRONG BUY";
  }
  
  // Branch 4: BUY
  if (ind.price > ind.sma200 && 
      ind.sma50 > ind.sma200 && 
      ind.rsi > 50 && ind.rsi <= 60 && 
      ind.macdHist > 0 && 
      ind.adx >= 15) {
    return "BUY";
  }
  
  // Branch 5: ACCUMULATE
  if (ind.price > ind.sma200 && 
      ind.rsi >= 35 && ind.rsi <= 55 && 
      ind.price >= ind.sma50 * 0.95 && 
      ind.price <= ind.sma50 * 1.05) {
    return "ACCUMULATE";
  }
  
  // Branch 6: DEEP VALUE
  if (ind.rsi <= 30 && 
      ind.price > ind.support &&
      (ind.fundamental === "VALUE" || ind.fundamental === "FAIR")) {
    return "DEEP VALUE";
  }
  
  // Branch 7: OVERSOLD WATCH
  if (ind.rsi <= 30 && ind.price > ind.support) {
    return "OVERSOLD WATCH";
  }
  
  // Branch 8: TRIM
  if (ind.rsi >= 70 || 
      ind.bollingerPctB >= 0.85 || 
      ind.price >= ind.resistance * 0.98) {
    return "TRIM";
  }
  
  // Branch 9: HOLD
  if (ind.price > ind.sma200 && 
      ind.rsi > 40 && ind.rsi < 70) {
    return "HOLD";
  }
  
  // Branch 10: DEFAULT
  return "NEUTRAL";
}

/**
 * Evaluates SIGNAL for TRADE mode
 * @param {object} ind - Indicator values
 * @param {number} atrAverage - 20-period ATR average
 * @private
 */
function evaluateSignalTrade(ind, atrAverage) {
  // Branch 1: STOP OUT (with 3% buffer)
  if (ind.price < ind.support * 0.97) {
    return "STOP OUT";
  }
  
  // Branch 2: VOLATILITY BREAKOUT
  if (atrAverage > 0 && ind.atr > atrAverage * 1.5 && 
      ind.volTrend >= 2.0 && ind.price >= ind.resistance * 1.01) {
    return "VOLATILITY BREAKOUT";
  }
  
  // Branch 3: BREAKOUT
  if (ind.volTrend >= 1.5 && ind.price >= ind.resistance * 1.02) {
    return "BREAKOUT";
  }
  
  // Branch 4: ATH BREAKOUT
  if (ind.athDiff >= -0.01 && ind.volTrend >= 2.0 && ind.adx >= 25) {
    return "ATH BREAKOUT";
  }
  
  // Branch 5: MOMENTUM
  if (ind.price > ind.sma50 && ind.macdHist > 0 && ind.adx >= 20) {
    return "MOMENTUM";
  }
  
  // Branch 6: OVERSOLD REVERSAL
  if (ind.stochK <= 20 && ind.macdHist > 0 && ind.price > ind.support) {
    return "OVERSOLD REVERSAL";
  }
  
  // Branch 7: VOLATILITY SQUEEZE
  if (atrAverage > 0 && ind.atr < atrAverage * 0.7 && 
      ind.adx < 15 && Math.abs(ind.bollingerPctB - 0.5) < 0.2) {
    return "VOLATILITY SQUEEZE";
  }
  
  // Branch 8: RANGE SUPPORT BUY
  if (ind.adx < 15 && 
      ind.price >= ind.support * 0.98 && 
      ind.price <= ind.support * 1.02) {
    return "RANGE SUPPORT BUY";
  }
  
  // Branch 9: OVERBOUGHT
  if (ind.rsi >= 70 || ind.bollingerPctB >= 0.9) {
    return "OVERBOUGHT";
  }
  
  // Branch 10: RISK OFF
  if (ind.price < ind.sma200) {
    return "RISK OFF";
  }
  
  // Branch 11: RANGE
  if (ind.adx < 15 && ind.price > ind.support) {
    return "RANGE";
  }
  
  // Branch 12: DEFAULT
  return "NEUTRAL";
}

/**
 * Evaluates DECISION for LONG-TERM INVESTMENT mode
 * @private
 */
function evaluateDecisionLongTerm(signal, isPurchased, hasPattern, patternType) {
  if (isPurchased) {
    if (signal === "STOP OUT" || signal === "RISK OFF") return "🔴 EXIT";
    if (signal === "TRIM" && hasPattern && patternType === "bearish") return "🟠 TRIM (PATTERN CONFIRMED)";
    if (signal === "TRIM") return "🟠 TRIM";
    if ((signal === "STRONG BUY" || signal === "BUY" || signal === "ACCUMULATE") && hasPattern && patternType === "bullish") return "🟢 ADD (PATTERN CONFIRMED)";
    if ((signal === "STRONG BUY" || signal === "BUY" || signal === "ACCUMULATE") && hasPattern && patternType === "bearish") return "⚠️ HOLD (PATTERN CONFLICT)";
    if (signal === "STRONG BUY" || signal === "BUY" || signal === "ACCUMULATE") return "🟢 ADD";
    if (signal === "HOLD") return "⚖️ HOLD";
    return "⚖️ HOLD";
  } else {
    if (signal === "STOP OUT" || signal === "RISK OFF") return "🔴 AVOID";
    if (signal === "STRONG BUY" && hasPattern && patternType === "bullish") return "🟢 STRONG BUY (PATTERN CONFIRMED)";
    if ((signal === "STRONG BUY" || signal === "BUY") && hasPattern && patternType === "bearish") return "⚠️ CAUTION (PATTERN CONFLICT)";
    if (signal === "STRONG BUY") return "🟢 STRONG BUY";
    if (signal === "BUY") return "🟢 BUY";
    if (signal === "ACCUMULATE") return "🟢 ACCUMULATE";
    if (signal === "DEEP VALUE") return "🟢 DEEP VALUE BUY";
    if (signal === "OVERSOLD WATCH") return "🟡 WATCH (OVERSOLD)";
    if (signal === "TRIM") return "⏳ WAIT (EXTENDED)";
    if (signal === "HOLD") return "⚖️ WATCH";
    return "⚪ NEUTRAL";
  }
}

/**
 * Evaluates DECISION for TRADE mode
 * @private
 */
function evaluateDecisionTrade(signal, isPurchased, hasPattern, patternType, price, support, resistance) {
  // Branch 1: STOP OUT check
  if (price > 0 && support > 0 && price < support * 0.97) {
    return "🔴 STOP OUT";
  }
  
  if (!isPurchased) {
    if (signal === "VOLATILITY BREAKOUT" && hasPattern && patternType === "bullish") return "🟢 STRONG TRADE LONG (PATTERN CONFIRMED)";
    if ((signal === "BREAKOUT" || signal === "ATH BREAKOUT") && hasPattern && patternType === "bullish") return "🟢 TRADE LONG (PATTERN CONFIRMED)";
    if ((signal === "VOLATILITY BREAKOUT" || signal === "BREAKOUT" || signal === "ATH BREAKOUT" || signal === "MOMENTUM") && hasPattern && patternType === "bearish") return "⚠️ CAUTION (PATTERN CONFLICT)";
    if (signal === "VOLATILITY BREAKOUT") return "🟢 STRONG TRADE LONG";
    if (signal === "BREAKOUT" || signal === "ATH BREAKOUT") return "🟢 TRADE LONG";
    if (signal === "MOMENTUM") return "🟡 ACCUMULATE";
    if (signal === "OVERSOLD REVERSAL") return "🟢 BUY DIP";
    if (signal === "RANGE SUPPORT BUY") return "🟡 RANGE BUY";
    if (signal === "VOLATILITY SQUEEZE") return "⏳ WAIT FOR BREAKOUT";
    if (signal === "RISK OFF") return "🔴 AVOID";
    return "⚪ NEUTRAL";
  } else {
    if (signal === "OVERBOUGHT" || (price > 0 && resistance > 0 && price >= resistance * 0.98)) return "🟠 TAKE PROFIT";
    if (signal === "RISK OFF") return "🔴 RISK OFF";
    return "⚖️ HOLD";
  }
}
