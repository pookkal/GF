/**
* ==============================================================================
* STABLE_MASTER_ALL_CLEAN_v3.1_KIRO_OPTIMIZED
* ==============================================================================
*/

// Delay constants for staggered formula writing (in milliseconds)
// These delays prevent calculation engine overload, especially on Android app
const DELAY_AFTER_MAIN_FORMULAS = 12500;  // 12.5 seconds - allows calculation engine to process bulk formulas (columns E-AF)
const DELAY_AFTER_CD_FORMULAS = 2000;     // 2 seconds - shorter delay for smaller formula set (columns C-D)

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

    // Ensure sheet has enough columns (32 total: A-AF)
    const maxCols = calc.getMaxColumns();
    if (maxCols < 32) {
      calc.insertColumnsAfter(maxCols, 32 - maxCols);
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
  styleGroup("B1:D1", "SIGNALING", COLORS.SIGNALING);
  styleGroup("E1:G1", "PRICE / VOLUME", COLORS.PRICE_VOLUME);
  styleGroup("H1:K1", "PERFORMANCE", COLORS.PERFORMANCE);  // H-K includes FUNDAMENTAL
  styleGroup("L1:O1", "TREND", COLORS.TREND);
  styleGroup("P1:T1", "MOMENTUM", COLORS.MOMENTUM);
  styleGroup("U1:X1", "VOLATILITY", COLORS.VOLATILITY);
  styleGroup("Y1:AE1", "TARGET", COLORS.TARGET);  // Y-AE all TARGET (no LEVELS)
  styleGroup("AF1:AF1", "NOTES", COLORS.NOTES);

  // Timestamp in AF1
  calc.getRange("AF1")
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
    COLORS.SIGNALING, COLORS.SIGNALING, COLORS.SIGNALING,  // B-D: SIGNAL, PATTERNS, DECISION
    COLORS.PRICE_VOLUME, COLORS.PRICE_VOLUME, COLORS.PRICE_VOLUME,  // E-G: Price, Change%, Vol Trend
    COLORS.PERFORMANCE, COLORS.PERFORMANCE, COLORS.PERFORMANCE, COLORS.PERFORMANCE,  // H-K: ATH TRUE, ATH Diff%, ATH ZONE, FUNDAMENTAL
    COLORS.TREND, COLORS.TREND, COLORS.TREND, COLORS.TREND,  // L-O: Trend State, SMA 20/50/200
    COLORS.MOMENTUM, COLORS.MOMENTUM, COLORS.MOMENTUM, COLORS.MOMENTUM, COLORS.MOMENTUM,  // P-T: RSI, MACD, Div, ADX, Stoch
    COLORS.VOLATILITY, COLORS.VOLATILITY, COLORS.VOLATILITY, COLORS.VOLATILITY,  // U-X: VOL REGIME, BBP SIGNAL, ATR, Bollinger %B
    COLORS.TARGET, COLORS.TARGET,  // Y-Z: Target, R:R Quality
    COLORS.TARGET, COLORS.TARGET, COLORS.TARGET, COLORS.TARGET, COLORS.TARGET,  // AA-AE: Support, Resistance, ATR STOP/TARGET, Position Size
    COLORS.NOTES  // AF: LAST STATE
  ];

  // Set Row 2 headers with group colors
  calc.getRange(2, 1, 1, 32)
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
  calc.setColumnWidth(2, 140);  // B: SIGNAL
  calc.setColumnWidth(3, 120);  // C: PATTERNS
  calc.setColumnWidth(4, 180);  // D: DECISION
  calc.setColumnWidth(5, 80);   // E: Price
  calc.setColumnWidth(6, 80);   // F: Change %
  calc.setColumnWidth(7, 80);   // G: Vol Trend
  calc.setColumnWidth(8, 90);   // H: ATH (TRUE)
  calc.setColumnWidth(9, 80);   // I: ATH Diff %
  calc.setColumnWidth(10, 120); // J: ATH ZONE
  calc.setColumnWidth(11, 140); // K: FUNDAMENTAL
  calc.setColumnWidth(12, 100); // L: Trend State
  calc.setColumnWidth(13, 80);  // M: SMA 20
  calc.setColumnWidth(14, 80);  // N: SMA 50
  calc.setColumnWidth(15, 80);  // O: SMA 200
  calc.setColumnWidth(16, 70);  // P: RSI
  calc.setColumnWidth(17, 80);  // Q: MACD Hist
  calc.setColumnWidth(18, 100); // R: Divergence
  calc.setColumnWidth(19, 70);  // S: ADX
  calc.setColumnWidth(20, 90);  // T: Stoch %K
  calc.setColumnWidth(21, 110); // U: VOL REGIME
  calc.setColumnWidth(22, 130); // V: BBP SIGNAL
  calc.setColumnWidth(23, 70);  // W: ATR
  calc.setColumnWidth(24, 100); // X: Bollinger %B
  calc.setColumnWidth(25, 80);  // Y: Target
  calc.setColumnWidth(26, 90);  // Z: R:R Quality
  calc.setColumnWidth(27, 80);  // AA: Support
  calc.setColumnWidth(28, 90);  // AB: Resistance
  calc.setColumnWidth(29, 90);  // AC: ATR STOP
  calc.setColumnWidth(30, 100); // AD: ATR TARGET
  calc.setColumnWidth(31, 120); // AE: POSITION SIZE
  calc.setColumnWidth(32, 120); // AF: LAST STATE

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
  const numCols = 32; // A-AF
  
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
    LIGHT_COLORS.SIGNALING, LIGHT_COLORS.SIGNALING, LIGHT_COLORS.SIGNALING,  // B-D
    LIGHT_COLORS.PRICE_VOLUME, LIGHT_COLORS.PRICE_VOLUME, LIGHT_COLORS.PRICE_VOLUME,  // E-G
    LIGHT_COLORS.PERFORMANCE, LIGHT_COLORS.PERFORMANCE, LIGHT_COLORS.PERFORMANCE, LIGHT_COLORS.PERFORMANCE,  // H-K
    LIGHT_COLORS.TREND, LIGHT_COLORS.TREND, LIGHT_COLORS.TREND, LIGHT_COLORS.TREND,  // L-O
    LIGHT_COLORS.MOMENTUM, LIGHT_COLORS.MOMENTUM, LIGHT_COLORS.MOMENTUM, LIGHT_COLORS.MOMENTUM, LIGHT_COLORS.MOMENTUM,  // P-T
    LIGHT_COLORS.VOLATILITY, LIGHT_COLORS.VOLATILITY, LIGHT_COLORS.VOLATILITY, LIGHT_COLORS.VOLATILITY,  // U-X
    LIGHT_COLORS.TARGET, LIGHT_COLORS.TARGET,  // Y-Z
    LIGHT_COLORS.TARGET, LIGHT_COLORS.TARGET, LIGHT_COLORS.TARGET, LIGHT_COLORS.TARGET, LIGHT_COLORS.TARGET,  // AA-AE
    LIGHT_COLORS.NOTES  // AF
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
  const inputSheet = ss.getSheetByName('INPUT');
  const useLongTermSignal = inputSheet.getRange('E2').getValue() === true;
  
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
  
  // STEP 2: Write Phase 1 formulas (columns E-AF) for ALL tickers at once
  Logger.log('Step 2: Writing Phase 1 formulas (columns E-AF) for all tickers...');
  ss.toast('Writing main formulas (E-AF)...', '⏳ Phase 1', 3);
  
  const phase1Data = [];
  for (let i = 0; i < allFormulas.length; i++) {
    if (allFormulas[i] && allFormulas[i].formulas) {
      phase1Data.push(allFormulas[i].formulas.slice(3)); // Indices 3-30 (28 columns)
    } else {
      phase1Data.push(new Array(28).fill('')); // Empty row for failed formulas
    }
  }
  
  try {
    if (phase1Data.length > 0) {
      calc.getRange(3, 5, phase1Data.length, 28).setFormulas(phase1Data);
      Logger.log(`Phase 1 complete: Wrote formulas for columns E-AF (${phase1Data.length} tickers)`);
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
  
  // STEP 3: Write Phase 2 formulas (columns C-D) for ALL tickers at once
  Logger.log('Step 3: Writing Phase 2 formulas (columns C-D) for all tickers...');
  ss.toast('Writing pattern formulas (C-D)...', '⏳ Phase 2', 3);
  
  const phase2Data = [];
  for (let i = 0; i < allFormulas.length; i++) {
    if (allFormulas[i] && allFormulas[i].formulas) {
      phase2Data.push(allFormulas[i].formulas.slice(1, 3)); // Indices 1-2 (2 columns)
    } else {
      phase2Data.push(['', '']); // Empty row for failed formulas
    }
  }
  
  try {
    if (phase2Data.length > 0) {
      calc.getRange(3, 3, phase2Data.length, 2).setFormulas(phase2Data);
      Logger.log(`Phase 2 complete: Wrote formulas for columns C-D (${phase2Data.length} tickers)`);
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
  ss.toast('Writing signal formulas (B)...', '⏳ Phase 3', 3);
  
  const phase3Data = [];
  for (let i = 0; i < allFormulas.length; i++) {
    if (allFormulas[i] && allFormulas[i].formulas) {
      phase3Data.push([allFormulas[i].formulas[0]]); // Index 0 (1 column)
    } else {
      phase3Data.push(['']); // Empty row for failed formulas
    }
  }
  
  try {
    if (phase3Data.length > 0) {
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
    // Apply percentage formatting to columns F, I, T, X for all data rows
    const numRows = tickers.length;
    calc.getRange(3, 6, numRows, 1).setNumberFormat('0.00%');  // F: Change %
    calc.getRange(3, 9, numRows, 1).setNumberFormat('0.00%');  // I: ATH Diff %
    calc.getRange(3, 20, numRows, 1).setNumberFormat('0.00%'); // T: Stoch %K
    calc.getRange(3, 24, numRows, 1).setNumberFormat('0.00%'); // X: Bollinger %B
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
    
    // Build all formulas for this ticker (31 formulas: B-AF)
    // CORRECT COLUMN ORDER per ColumnMapping.js CALC_COLUMNS:
    // B=SIGNAL, C=PATTERNS, D=DECISION, E=Price, F=Change%, G=Vol Trend,
    // H=ATH TRUE, I=ATH Diff%, J=ATH ZONE, K=FUNDAMENTAL,
    // L=Trend State, M=SMA 20, N=SMA 50, O=SMA 200,
    // P=RSI, Q=MACD Hist, R=Divergence, S=ADX, T=Stoch %K,
    // U=VOL REGIME, V=BBP SIGNAL, W=ATR, X=Bollinger %B,
    // Y=Target, Z=R:R Quality, AA=Support, AB=Resistance, AC=ATR STOP, AD=ATR TARGET,
    // AE=POSITION SIZE, AF=LAST STATE
    
    const formulas = [
      buildSignalFormula(row, SEP, useLongTermSignal),                    // B: SIGNAL
      `=IF($A${row}=""${SEP}""${SEP}GETPATTERNS($A${row}${SEP}$E${row}))`, // C: PATTERNS
      buildDecisionFormula(row, SEP, useLongTermSignal),                  // D: DECISION (uses B+C only)
      `=ROUND(IFERROR(GOOGLEFINANCE("${t}"${SEP}"price")${SEP}0)${SEP}2)`, // E: Price
      `=IFERROR(GOOGLEFINANCE("${t}"${SEP}"changepct")/100${SEP}0)`,      // F: Change %
      buildRVOLFormula(row, volCol, lastRowCount, SEP),                   // G: Vol Trend
      `=IFERROR(${athCell}${SEP}0)`,                                      // H: ATH (TRUE) - reads from DATA sheet
      `=IFERROR(($E${row}-$H${row})/MAX(0.01${SEP}$H${row})${SEP}0)`,    // I: ATH Diff % - uses H not I!
      buildATHZoneFormula(row, SEP),                                      // J: ATH ZONE - uses I not J!
      buildFundamentalFormula(row, peCell, epsCell, SEP),                 // K: FUNDAMENTAL - uses I not J!
      `=IF($E${row}>$O${row}${SEP}"BULL"${SEP}"BEAR")`,                   // L: Trend State - uses O (SMA 200) not P!
      buildSMAFormula(closeCol, lastRowCount, 20, SEP),                   // M: SMA 20
      buildSMAFormula(closeCol, lastRowCount, 50, SEP),                   // N: SMA 50
      buildSMAFormula(closeCol, lastRowCount, 200, SEP),                  // O: SMA 200
      `=LIVERSI(DATA!${closeCol}$5:${closeCol}${SEP}$E${row})`,          // P: RSI
      `=LIVEMACD(DATA!${closeCol}$5:${closeCol}${SEP}$E${row})`,         // Q: MACD Hist
      buildDivergenceFormula(row, closeCol, lastAbsRow, SEP),             // R: Divergence - uses Q not R!
      `=IFERROR(LIVEADX(DATA!${highCol}$5:${highCol}${SEP}DATA!${lowCol}$5:${lowCol}${SEP}DATA!${closeCol}$5:${closeCol}${SEP}$E${row})${SEP}0)`, // S: ADX
      `=LIVESTOCHK(DATA!${highCol}$5:${highCol}${SEP}DATA!${lowCol}$5:${lowCol}${SEP}DATA!${closeCol}$5:${closeCol}${SEP}$E${row})`, // T: Stoch %K
      buildVolRegimeFormula(row, SEP),                                    // U: VOL REGIME - moved from H!
      buildBBPSignalFormula(row, SEP),                                    // V: BBP SIGNAL - uses P and W not Q and X!
      `=IFERROR(LIVEATR(DATA!${highCol}$5:${highCol}${SEP}DATA!${lowCol}$5:${lowCol}${SEP}DATA!${closeCol}$5:${closeCol}${SEP}$E${row})${SEP}0)`, // W: ATR
      buildBBPFormula(row, closeCol, lastRowCount, SEP),                  // X: Bollinger %B - uses N (SMA 20) not O!
      `=ROUND(MAX($AB${row}${SEP}$E${row}+(($E${row}-$AA${row})*3))${SEP}2)`, // Y: Target
      buildRRFormula(row, SEP),                                           // Z: R:R Quality
      buildSupportFormula(row, lowCol, lastRowCount, SEP),                // AA: Support - uses S (ADX) not U!
      buildResistanceFormula(row, highCol, lastRowCount, SEP),            // AB: Resistance - uses S (ADX) not U!
      `=ROUND(MAX($AA${row}${SEP}$E${row}-($W${row}*2))${SEP}2)`,        // AC: ATR STOP
      `=ROUND($E${row}+($W${row}*3)${SEP}2)`,                             // AD: ATR TARGET
      buildPositionSizeFormula(row, SEP),                                 // AE: POSITION SIZE - uses I and Z not J and AA!
      `=IF($A${row}=""${SEP}""${SEP}$D${row})`                            // AF: LAST STATE
    ];
    
    // Validate that we have exactly 31 formulas
    if (formulas.length !== 31) {
      throw new Error(`Formula count mismatch: expected 31, got ${formulas.length}`);
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
  // CORRECT Column references per ColumnMapping.js CALC_COLUMNS:
  // E=Price, G=Vol Trend, H=ATH TRUE, I=ATH Diff%, M=SMA 20, N=SMA 50, O=SMA 200,
  // P=RSI, Q=MACD Hist, S=ADX, T=Stoch %K, W=ATR, X=Bollinger %B, AA=Support, AB=Resistance
  
  if (useLongTermSignal) {
    return `=IF(OR(ISBLANK($E${row})${SEP}$E${row}=0)${SEP}"LOADING"${SEP}IFS($E${row}<$AA${row}${SEP}"STOP OUT"${SEP}$E${row}<$O${row}${SEP}"RISK OFF"${SEP}AND($I${row}>=-0.01${SEP}$G${row}>=1.5${SEP}$S${row}>=20${SEP}$E${row}>$O${row})${SEP}"ATH BREAKOUT"${SEP}AND($W${row}>IFERROR(AVERAGE(OFFSET($W${row}${SEP}-MIN(20${SEP}ROW($W${row})-1)${SEP}0${SEP}MIN(20${SEP}ROW($W${row})-1)))${SEP}$W${row})*1.5${SEP}$G${row}>=2.0${SEP}$E${row}>$AB${row})${SEP}"VOLATILITY BREAKOUT"${SEP}AND($X${row}<=0.1${SEP}$P${row}<=25${SEP}$T${row}<=0.20${SEP}$E${row}>$O${row})${SEP}"EXTREME OVERSOLD BUY"${SEP}AND($E${row}>$O${row}${SEP}$N${row}>$O${row}${SEP}$P${row}<=30${SEP}$Q${row}>0${SEP}$S${row}>=20${SEP}$G${row}>=1.5)${SEP}"STRONG BUY"${SEP}AND($E${row}>$O${row}${SEP}$N${row}>$O${row}${SEP}$P${row}<=40${SEP}$Q${row}>0${SEP}$S${row}>=15)${SEP}"BUY"${SEP}AND($E${row}>$O${row}${SEP}$P${row}<=35${SEP}$E${row}>=$N${row}*0.95)${SEP}"ACCUMULATE"${SEP}$P${row}<=20${SEP}"OVERSOLD"${SEP}OR($P${row}>=80${SEP}$X${row}>=0.9)${SEP}"OVERBOUGHT"${SEP}AND($E${row}>$O${row}${SEP}$P${row}>40${SEP}$P${row}<70)${SEP}"HOLD"${SEP}TRUE${SEP}"NEUTRAL"))`;
  } else {
    return `=IF(OR(ISBLANK($E${row})${SEP}$E${row}=0)${SEP}"LOADING"${SEP}IFS($E${row}<$AA${row}${SEP}"STOP OUT"${SEP}$E${row}<$O${row}${SEP}"RISK OFF"${SEP}AND($W${row}>IFERROR(AVERAGE(OFFSET($W${row}${SEP}-MIN(20${SEP}ROW($W${row})-1)${SEP}0${SEP}MIN(20${SEP}ROW($W${row})-1)))${SEP}$W${row})*1.5${SEP}$G${row}>=2.0${SEP}$E${row}>$AB${row})${SEP}"VOLATILITY BREAKOUT"${SEP}AND($I${row}>=-0.01${SEP}$G${row}>=1.5${SEP}$S${row}>=20)${SEP}"ATH BREAKOUT"${SEP}AND($G${row}>=1.5${SEP}$E${row}>=$AB${row}*0.995)${SEP}"BREAKOUT"${SEP}AND($E${row}>$O${row}${SEP}$Q${row}>0${SEP}$S${row}>=20)${SEP}"MOMENTUM"${SEP}AND($E${row}>$O${row}${SEP}$N${row}>$O${row}${SEP}$S${row}>=15)${SEP}"UPTREND"${SEP}AND($E${row}>$N${row}${SEP}$E${row}>$M${row})${SEP}"BULLISH"${SEP}AND(OR($T${row}<=0.20${SEP}$X${row}<=0.2)${SEP}$E${row}>$AA${row})${SEP}"OVERSOLD"${SEP}OR($P${row}>=80${SEP}$X${row}>=0.9)${SEP}"OVERBOUGHT"${SEP}AND($W${row}<IFERROR(AVERAGE(OFFSET($W${row}${SEP}-MIN(20${SEP}ROW($W${row})-1)${SEP}0${SEP}MIN(20${SEP}ROW($W${row})-1)))${SEP}$W${row})*0.7${SEP}$S${row}<15${SEP}ABS($X${row}-0.5)<0.2)${SEP}"VOLATILITY SQUEEZE"${SEP}$S${row}<15${SEP}"RANGE"${SEP}TRUE${SEP}"NEUTRAL"))`;
  }
}

function buildFundamentalFormula(row, peCell, epsCell, SEP) {
  // ATH Diff % is in column I (not J!)
  return `=IFERROR(LET(peRaw${SEP}${peCell}${SEP}epsRaw${SEP}${epsCell}${SEP}athDiffRaw${SEP}$I${row}${SEP}pe${SEP}IFERROR(VALUE(REGEXREPLACE(TO_TEXT(peRaw)${SEP}"[^0-9\\.\\-]"${SEP}""))${SEP}"")${SEP}eps${SEP}IFERROR(VALUE(REGEXREPLACE(TO_TEXT(epsRaw)${SEP}"[^0-9\\.\\-]"${SEP}""))${SEP}"")${SEP}athDiff${SEP}IFERROR(VALUE(REGEXREPLACE(TO_TEXT(athDiffRaw)${SEP}"[^0-9\\.\\-]"${SEP}""))/100${SEP}"")${SEP}IFS(OR(pe=""${SEP}eps="")${SEP}"FAIR"${SEP}eps<=0${SEP}"ZOMBIE"${SEP}AND(pe>=60${SEP}athDiff<>""${SEP}athDiff>=-0.08)${SEP}"PRICED FOR PERFECTION"${SEP}pe>=35${SEP}"EXPENSIVE"${SEP}AND(pe>0${SEP}pe<=25${SEP}eps>=0.5)${SEP}"VALUE"${SEP}AND(pe>25${SEP}pe<35${SEP}eps>=0.5)${SEP}"FAIR"${SEP}TRUE${SEP}"FAIR"))${SEP}"FAIR")`;
}

function buildDecisionFormula(row, SEP, useLongTermSignal) {
  // CRITICAL: DECISION uses SIGNAL (B) + PATTERNS (C) only
  // FUNDAMENTAL (L) is informational but does NOT drive DECISION logic
  
  const tagExpr = `UPPER(IFERROR(INDEX(INPUT!$C$3:$C${SEP}MATCH($A${row}${SEP}INPUT!$A$3:$A${SEP}0))${SEP}""))`;
  const purchasedExpr = `ISNUMBER(SEARCH("PURCHASED"${SEP}${tagExpr}))`;
  
  // Pattern analysis helpers - Updated to use short forms
  const hasBullishPattern = `OR(ISNUMBER(SEARCH("ASC_TRI"${SEP}$C${row}))${SEP}ISNUMBER(SEARCH("BRKOUT"${SEP}$C${row}))${SEP}ISNUMBER(SEARCH("DBL_BTM"${SEP}$C${row}))${SEP}ISNUMBER(SEARCH("INV_H&S"${SEP}$C${row}))${SEP}ISNUMBER(SEARCH("CUP_HDL"${SEP}$C${row})))`;
  const hasBearishPattern = `OR(ISNUMBER(SEARCH("DESC_TRI"${SEP}$C${row}))${SEP}ISNUMBER(SEARCH("H&S"${SEP}$C${row}))${SEP}ISNUMBER(SEARCH("DBL_TOP"${SEP}$C${row})))`;
  const hasPattern = `NOT(OR($C${row}=""${SEP}$C${row}="—"))`;
  
  if (useLongTermSignal) {
    // Long-term investment mode: SIGNAL + PATTERNS (no FUNDAMENTAL in logic)
    return `=IF($A${row}=""${SEP}""${SEP}IF($B${row}="LOADING"${SEP}"LOADING"${SEP}IF(${purchasedExpr}${SEP}` +
      // For PURCHASED positions
      `IFS(` +
      `OR($B${row}="STOP OUT"${SEP}$B${row}="RISK OFF")${SEP}"🔴 EXIT"${SEP}` +
      `AND(OR($B${row}="STRONG BUY"${SEP}$B${row}="BUY"${SEP}$B${row}="ACCUMULATE")${SEP}${hasPattern}${SEP}${hasBullishPattern})${SEP}"🟢 ADD (PATTERN CONFIRMED)"${SEP}` +
      `AND(OR($B${row}="STRONG BUY"${SEP}$B${row}="BUY"${SEP}$B${row}="ACCUMULATE")${SEP}${hasPattern}${SEP}${hasBearishPattern})${SEP}"⚠️ HOLD (PATTERN CONFLICT)"${SEP}` +
      `OR($B${row}="STRONG BUY"${SEP}$B${row}="BUY"${SEP}$B${row}="ACCUMULATE")${SEP}"🟢 ADD"${SEP}` +
      `$B${row}="OVERBOUGHT"${SEP}"🟠 TRIM"${SEP}` +
      `$B${row}="HOLD"${SEP}"⚖️ HOLD"${SEP}` +
      `TRUE${SEP}"⚖️ HOLD"` +
      `)${SEP}` +
      // For NON-PURCHASED positions
      `IFS(` +
      `OR($B${row}="STOP OUT"${SEP}$B${row}="RISK OFF")${SEP}"🔴 AVOID"${SEP}` +
      `AND($B${row}="STRONG BUY"${SEP}${hasPattern}${SEP}${hasBullishPattern})${SEP}"🟢 STRONG BUY (PATTERN CONFIRMED)"${SEP}` +
      `AND(OR($B${row}="STRONG BUY"${SEP}$B${row}="BUY")${SEP}${hasPattern}${SEP}${hasBearishPattern})${SEP}"⚠️ HOLD (PATTERN CONFLICT)"${SEP}` +
      `$B${row}="STRONG BUY"${SEP}"🟢 STRONG BUY"${SEP}` +
      `OR($B${row}="BUY"${SEP}$B${row}="ACCUMULATE")${SEP}"🟢 BUY"${SEP}` +
      `$B${row}="OVERSOLD"${SEP}"🟡 WATCH (OVERSOLD)"${SEP}` +
      `$B${row}="OVERBOUGHT"${SEP}"⏳ WAIT (OVERBOUGHT)"${SEP}` +
      `$B${row}="HOLD"${SEP}"⚖️ WATCH"${SEP}` +
      `TRUE${SEP}"⚪ NEUTRAL"` +
      `)` +
      `)))`;
  } else {
    // Trade mode: SIGNAL + PATTERNS (no FUNDAMENTAL in logic)
    return `=IF($A${row}=""${SEP}""${SEP}LET(` +
      `tag${SEP}UPPER(IFERROR(INDEX(INPUT!$C$3:$C${SEP}MATCH($A${row}${SEP}INPUT!$A$3:$A${SEP}0))${SEP}""))${SEP}` +
      `purchased${SEP}REGEXMATCH(tag${SEP}"(^|,|\\\\s)PURCHASED(\\\\s|,|$)")${SEP}` +
      `bullishPat${SEP}${hasBullishPattern}${SEP}` +
      `bearishPat${SEP}${hasBearishPattern}${SEP}` +
      `hasPat${SEP}${hasPattern}${SEP}` +
      `IFS(` +
      // Stop-out check - Price below Support (AA not AB)
      `AND(IFERROR(VALUE($E${row})${SEP}0)>0${SEP}IFERROR(VALUE($AA${row})${SEP}0)>0${SEP}IFERROR(VALUE($E${row})${SEP}0)<IFERROR(VALUE($AA${row})${SEP}0))${SEP}"Stop-Out"${SEP}` +
      // Pattern-confirmed strong signals
      `AND(NOT(purchased)${SEP}OR($B${row}="VOLATILITY BREAKOUT"${SEP}$B${row}="ATH BREAKOUT")${SEP}hasPat${SEP}bullishPat)${SEP}"🟢 STRONG TRADE LONG (PATTERN CONFIRMED)"${SEP}` +
      `AND(NOT(purchased)${SEP}$B${row}="BREAKOUT"${SEP}hasPat${SEP}bullishPat)${SEP}"🟢 TRADE LONG (PATTERN CONFIRMED)"${SEP}` +
      // Pattern conflicts
      `AND(NOT(purchased)${SEP}OR($B${row}="VOLATILITY BREAKOUT"${SEP}$B${row}="ATH BREAKOUT"${SEP}$B${row}="BREAKOUT"${SEP}$B${row}="MOMENTUM")${SEP}hasPat${SEP}bearishPat)${SEP}"⚠️ HOLD (PATTERN CONFLICT)"${SEP}` +
      // Standard signals without pattern consideration
      `AND(NOT(purchased)${SEP}OR($B${row}="VOLATILITY BREAKOUT"${SEP}$B${row}="ATH BREAKOUT"))${SEP}"Strong Trade Long"${SEP}` +
      `AND(NOT(purchased)${SEP}$B${row}="BREAKOUT")${SEP}"Trade Long"${SEP}` +
      `AND(NOT(purchased)${SEP}$B${row}="MOMENTUM")${SEP}"Accumulate"${SEP}` +
      `AND(NOT(purchased)${SEP}$B${row}="OVERSOLD")${SEP}"Add in Dip"${SEP}` +
      `AND(NOT(purchased)${SEP}$B${row}="VOLATILITY SQUEEZE")${SEP}"Wait for Breakout"${SEP}` +
      // Purchased position management
      `AND(purchased${SEP}OR($B${row}="OVERBOUGHT"${SEP}IFERROR(VALUE($E${row})${SEP}0)>=IFERROR(VALUE($AD${row})${SEP}0)))${SEP}"Take Profit"${SEP}` +
      `AND(purchased${SEP}$B${row}="RISK OFF")${SEP}"Risk-Off"${SEP}` +
      `AND(NOT(purchased)${SEP}$B${row}="RISK OFF")${SEP}"Avoid"${SEP}` +
      // Default holds
      `$B${row}="MOMENTUM"${SEP}"Hold"${SEP}` +
      `$B${row}="UPTREND"${SEP}"Hold"${SEP}` +
      `$B${row}="BULLISH"${SEP}"Hold"${SEP}` +
      `TRUE${SEP}"Hold"` +
      `)))`;
  }
}

function buildRVOLFormula(row, volCol, lastRowCount, SEP) {
  return `=ROUND(IFERROR(OFFSET(DATA!${volCol}$5${SEP}${lastRowCount}-1${SEP}0)/AVERAGE(OFFSET(DATA!${volCol}$5${SEP}${lastRowCount}-20${SEP}0${SEP}20))${SEP}1)${SEP}2)`;
}

function buildRRFormula(row, SEP) {
  // E=Price, W=ATR, AA=Support, AB=Resistance
  return `=IF(OR($E${row}<=$AA${row}${SEP}$E${row}=0)${SEP}0${SEP}ROUND(MAX(0${SEP}$AB${row}-$E${row})/MAX($W${row}*0.5${SEP}$E${row}-$AA${row})${SEP}2))`;
}

function buildSMAFormula(closeCol, lastRowCount, period, SEP) {
  return `=ROUND(IFERROR(AVERAGE(OFFSET(DATA!${closeCol}$5${SEP}${lastRowCount}-${period}${SEP}0${SEP}${period}))${SEP}0)${SEP}2)`;
}

function buildDivergenceFormula(row, closeCol, lastAbsRow, SEP) {
  // E=Price, P=RSI (not R!)
  return `=IFERROR(IFS(AND($E${row}<INDEX(DATA!${closeCol}:${closeCol}${SEP}${lastAbsRow}-14)${SEP}$P${row}>50)${SEP}"BULL DIV"${SEP}AND($E${row}>INDEX(DATA!${closeCol}:${closeCol}${SEP}${lastAbsRow}-14)${SEP}$P${row}<50)${SEP}"BEAR DIV"${SEP}TRUE${SEP}"—")${SEP}"—")`;
}

function buildSupportFormula(row, lowCol, lastRowCount, SEP) {
  // S=ADX (not U!)
  return `=ROUND(IFERROR(LET(win${SEP}IFS($S${row}<20${SEP}10${SEP}$S${row}<35${SEP}22${SEP}TRUE${SEP}40)${SEP}n${SEP}${lastRowCount}${SEP}start${SEP}MAX(0${SEP}n-win)${SEP}len${SEP}MIN(win${SEP}n)${SEP}rng${SEP}IF(len<=0${SEP}OFFSET(DATA!${lowCol}$5${SEP}0${SEP}0)${SEP}OFFSET(DATA!${lowCol}$5${SEP}start${SEP}0${SEP}len))${SEP}out${SEP}IF(COUNTA(rng)<3${SEP}IFERROR(MIN(rng)${SEP}0)${SEP}PERCENTILE.INC(rng${SEP}0.15))${SEP}out)${SEP}0)${SEP}2)`;
}

function buildResistanceFormula(row, highCol, lastRowCount, SEP) {
  // S=ADX (not U!)
  return `=ROUND(IFERROR(LET(win${SEP}IFS($S${row}<20${SEP}10${SEP}$S${row}<35${SEP}22${SEP}TRUE${SEP}40)${SEP}n${SEP}${lastRowCount}${SEP}start${SEP}MAX(0${SEP}n-win)${SEP}len${SEP}MIN(win${SEP}n)${SEP}rng${SEP}IF(len<=0${SEP}OFFSET(DATA!${highCol}$5${SEP}0${SEP}0)${SEP}OFFSET(DATA!${highCol}$5${SEP}start${SEP}0${SEP}len))${SEP}out${SEP}IF(COUNTA(rng)<3${SEP}IFERROR(MAX(rng)${SEP}0)${SEP}PERCENTILE.INC(rng${SEP}0.85))${SEP}out)${SEP}0)${SEP}2)`;
}

function buildBBPFormula(row, closeCol, lastRowCount, SEP) {
  // E=Price, M=SMA 20 (not O!)
  return `=ROUND(IFERROR((($E${row}-$M${row})/(4*STDEV(OFFSET(DATA!${closeCol}$5${SEP}${lastRowCount}-20${SEP}0${SEP}20))))+0.5${SEP}0.5)${SEP}2)`;
}

function buildPositionSizeFormula(row, SEP) {
  // E=Price, I=ATH Diff% (not J!), W=ATR, Z=R:R Quality
  return `=IF($A${row}=""${SEP}""${SEP}LET(riskReward${SEP}$Z${row}${SEP}atrRisk${SEP}$W${row}/$E${row}${SEP}athRisk${SEP}IF($I${row}>=-0.05${SEP}0.8${SEP}1.0)${SEP}volRegimeRisk${SEP}IFS(atrRisk<=0.02${SEP}1.2${SEP}atrRisk<=0.05${SEP}1.0${SEP}atrRisk<=0.08${SEP}0.7${SEP}TRUE${SEP}0.5)${SEP}baseSize${SEP}0.02${SEP}rrMultiplier${SEP}IF(riskReward>=3${SEP}1.5${SEP}IF(riskReward>=2${SEP}1.0${SEP}0.5))${SEP}finalSize${SEP}MIN(0.08${SEP}baseSize*rrMultiplier*volRegimeRisk*athRisk)${SEP}TEXT(finalSize${SEP}"0.0%")&" (Vol: "&IFS(atrRisk<=0.02${SEP}"LOW"${SEP}atrRisk<=0.05${SEP}"NORM"${SEP}atrRisk<=0.08${SEP}"HIGH"${SEP}TRUE${SEP}"EXTR")&")"))`;
}

function buildVolRegimeFormula(row, SEP) {
  // W=ATR, E=Price
  return `=IFS($W${row}/$E${row}<=0.02${SEP}"LOW VOL"${SEP}$W${row}/$E${row}<=0.05${SEP}"NORMAL VOL"${SEP}$W${row}/$E${row}<=0.08${SEP}"HIGH VOL"${SEP}TRUE${SEP}"EXTREME VOL")`;
}

function buildATHZoneFormula(row, SEP) {
  // I=ATH Diff % (not J!)
  return `=IFS($I${row}>=-0.02${SEP}"AT ATH"${SEP}$I${row}>=-0.05${SEP}"NEAR ATH"${SEP}$I${row}>=-0.15${SEP}"RESISTANCE ZONE"${SEP}$I${row}>=-0.30${SEP}"PULLBACK ZONE"${SEP}$I${row}>=-0.50${SEP}"CORRECTION ZONE"${SEP}TRUE${SEP}"DEEP VALUE ZONE")`;
}

function buildBBPSignalFormula(row, SEP) {
  // P=RSI (not Q!), X=Bollinger %B, E=Price, O=SMA 200 (not P!), AA=Support
  return `=IFS(AND($X${row}>=0.9${SEP}$P${row}>=70)${SEP}"EXTREME OVERBOUGHT"${SEP}AND($X${row}<=0.1${SEP}$P${row}<=30)${SEP}"EXTREME OVERSOLD"${SEP}AND($X${row}>=0.8${SEP}$E${row}>$O${row})${SEP}"MOMENTUM STRONG"${SEP}AND($X${row}<=0.2${SEP}$E${row}>$AA${row})${SEP}"MEAN REVERSION"${SEP}TRUE${SEP}"NEUTRAL")`;
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
    buildBBPSignalFormula
  };
}