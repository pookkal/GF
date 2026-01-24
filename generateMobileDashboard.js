/**
* ==============================================================================
* STABLE_MASTER_ALL_CLEAN_v3.1_KIRO_volume_fix
* ==============================================================================
*/


function generateMobileReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const REPORT = ss.getSheetByName('REPORT') || ss.insertSheet('REPORT');
  const INPUT = ss.getSheetByName('INPUT');
  
  if (!INPUT) throw new Error('INPUT sheet not found');
  
  // Setup dropdown
  setupReportTickerDropdown_(REPORT, INPUT);
  
  // Safely clear the sheet of all content and formatting
  safeClearSheet_(REPORT);
  
  // Apply column widths
  setReportColumnWidthsAndWrap___(REPORT);
  
  // Create report
  createFormulaReport_(REPORT);
  
  SpreadsheetApp.flush();
  
  // NOTE: Chart creation is now automatic on sheet generation
}

/**
 * Safely clear sheet and all merged ranges
 */
function safeClearSheet_(REPORT) {
  // Remove all charts first
  REPORT.getCharts().forEach(chart => REPORT.removeChart(chart));
  
  // Clear content and formats
  REPORT.clear();
  
  // Force unmerge all ranges by getting each merged range and breaking it apart
  let attempts = 0;
  const maxAttempts = 10;
  
  while (attempts < maxAttempts) {
    try {
      const mergedRanges = REPORT.getMergedRanges();
      if (mergedRanges.length === 0) break;
      
      // Unmerge each range using its exact range
      for (const range of mergedRanges) {
        try {
          range.breakApart();
        } catch (e) {
          // If individual range fails, continue
          console.log(`Failed to unmerge range: ${range.getA1Notation()}`);
        }
      }
      attempts++;
    } catch (e) {
      break;
    }
  }
}

/**
 * Create the complete formula-based report
 */
function createFormulaReport_(REPORT) {
  const P = reportPalette___();
  
  // Get locale separator
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const locale = (ss.getSpreadsheetLocale() || "").toLowerCase();
  const SEP = (/^(en|en_)/.test(locale)) ? "," : ";";
  
  // Set professional font for entire sheet
  const maxRows = Math.max(50, REPORT.getLastRow());
  REPORT.getRange(1, 1, maxRows, 12).setFontFamily('Calibri');
  
  // Helper function for robust lookups
  const lookup = (col) => `=IFERROR(INDEX(CALCULATIONS!${col}:${col}${SEP}MATCH(UPPER(TRIM($A$1))${SEP}ARRAYFORMULA(UPPER(TRIM(CALCULATIONS!A:A)))${SEP}0))${SEP}"—")`;
  
  // Numeric lookup for calculations
  const numLookup = (col) => `IFERROR(VALUE(INDEX(CALCULATIONS!${col}:${col}${SEP}MATCH(UPPER(TRIM($A$1))${SEP}ARRAYFORMULA(UPPER(TRIM(CALCULATIONS!A:A)))${SEP}0)))${SEP}0)`;
  
  // Row 1: Ticker name merged A1:C1 - Clean professional header
  REPORT.getRange('A1:C1').merge()
    .setBackground('#1E3A8A')  // Deep blue for ticker
    .setFontColor('#FFFFFF')
    .setFontWeight('normal')
    .setFontSize(16)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  
  // Set ticker dropdown validation on the merged cell A1:C1
  const INPUT = ss.getSheetByName('INPUT');
  if (INPUT) {
    setupReportTickerDropdown_(REPORT, INPUT);
  }
  
  // Row 2: Date selection dropdowns A2:C2
  // Years dropdown (A2): 0Y to 20Y
  const yearsValues = [];
  for (let i = 0; i <= 20; i++) {
    yearsValues.push([i + 'Y']);
  }
  REPORT.getRange('A2')
    .setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(yearsValues.flat(), true).build())
    .setValue('0Y')
    .setBackground('#374151')
    .setFontColor('#FFFFFF')
    .setFontWeight('normal')
    .setFontSize(10)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  
  // Months dropdown (B2): 0M to 12M
  const monthsValues = [];
  for (let i = 0; i <= 12; i++) {
    monthsValues.push([i + 'M']);
  }
  REPORT.getRange('B2')
    .setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(monthsValues.flat(), true).build())
    .setValue('1M')
    .setBackground('#374151')
    .setFontColor('#FFFFFF')
    .setFontWeight('normal')
    .setFontSize(10)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  
  // Days dropdown (C2): 0D to 31D
  const daysValues = [];
  for (let i = 0; i <= 31; i++) {
    daysValues.push([i + 'D']);
  }
  REPORT.getRange('C2')
    .setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(daysValues.flat(), true).build())
    .setValue('0D')
    .setBackground('#374151')
    .setFontColor('#FFFFFF')
    .setFontWeight('normal')
    .setFontSize(10)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  
  // Row 3: Date display (A3:B3 merged) and Interval dropdown (C3)
  // Calculated date display (A3:B3 merged) - Clean display
  REPORT.getRange('A3:B3').merge()
    .setFormula('=IF(AND(ISNUMBER(VALUE(LEFT(A2,LEN(A2)-1))),ISNUMBER(VALUE(LEFT(B2,LEN(B2)-1))),ISNUMBER(VALUE(LEFT(C2,LEN(C2)-1)))),TEXT(TODAY()-VALUE(LEFT(A2,LEN(A2)-1))*365-VALUE(LEFT(B2,LEN(B2)-1))*30-VALUE(LEFT(C2,LEN(C2)-1)),"yyyy-mm-dd"),"Select Date")')
    .setBackground('#1F2937')
    .setFontColor('#60A5FA')  // Light blue for date
    .setFontWeight('normal')
    .setFontSize(11)
    .setFontFamily('Calibri')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  
  // Interval dropdown C3 (Weekly/Daily)
  REPORT.getRange('C3')
    .setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Weekly', 'Daily'], true).build())
    .setValue('Daily')
    .setBackground('#374151')
    .setFontColor('#FFFFFF')
    .setFontWeight('normal')
    .setFontSize(10)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  
  // Chart controls starting from D1:M2 (labels in row 1, checkboxes in row 2)
  setupChartControls_(REPORT);
  
  // Decision section starts at row 4 (moved down to accommodate new row 3)
  // Order: DECISION, SIGNAL, PATTERNS
  REPORT.getRange('A4').setValue('DECISION');
  REPORT.getRange('B4').setFormula(lookup('C')); // DECISION from CALCULATIONS column C
  REPORT.getRange('B4:C4').merge();
  
  REPORT.getRange('A5').setValue('SIGNAL');
  REPORT.getRange('B5').setFormula(lookup('D')); // SIGNAL from CALCULATIONS column D
  REPORT.getRange('B5:C5').merge();
  
  REPORT.getRange('A6').setValue('PATTERNS');
  REPORT.getRange('B6').setFormula(lookup('E')); // PATTERNS from CALCULATIONS column E
  REPORT.getRange('B6:C6').merge();
  
  // Row 7: M.RATING (Market Rating from INPUT column D)
  REPORT.getRange('A7').setValue('M.RATING');
  REPORT.getRange('B7:C7').merge();
  REPORT.getRange('B7').setFormula(`=IFERROR(INDEX(INPUT!D:D${SEP}MATCH(UPPER(TRIM($A$1))${SEP}ARRAYFORMULA(UPPER(TRIM(INPUT!A:A)))${SEP}0))${SEP}"—")`);
  
  // Row 8: M.PRICE (Market Price from INPUT column E)
  REPORT.getRange('A8').setValue('M.PRICE');
  REPORT.getRange('B8:C8').merge();
  REPORT.getRange('B8').setFormula(`=IFERROR(INDEX(INPUT!E:E${SEP}MATCH(UPPER(TRIM($A$1))${SEP}ARRAYFORMULA(UPPER(TRIM(INPUT!A:A)))${SEP}0))${SEP}"—")`);
  
  // Style decision section (updated to include rows 7-8) - Professional styling with minimal borders
  REPORT.getRange('A4:C8')
    .setFontWeight('normal')
    .setFontSize(11)
    .setFontFamily('Calibri');
  
  // Add only outer border for clean look - WHITE borders
  REPORT.getRange('A4:C8')
    .setBorder(true, true, true, true, false, false, '#FFFFFF', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  
  // Set decision label cells (A4:A8) with professional gradient-like appearance
  REPORT.getRange('A4:A8')
    .setBackground('#2D3748')
    .setFontColor('#F59E0B')  // Amber color for labels
    .setFontWeight('normal')
    .setHorizontalAlignment('right')
    .setVerticalAlignment('middle');
  
  // Set decision value cells with clean dark background
  REPORT.getRange('B4:C8')
    .setBackground(P.BG_ROW_A)
    .setFontColor(P.TEXT)
    .setFontWeight('normal')
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle');
  
  // Apply conditional formatting to decision cells
  applyDecisionConditionalFormatting_(REPORT);
  
  // Apply conditional formatting to M.PRICE (B8) - green if M.PRICE > Current Price, red otherwise
  applyMarketPriceConditionalFormatting_(REPORT);
  
  // Chart section at D3:N17 - starts after controls
  setupChartSection_(REPORT);
  
  // Data rows start at row 9 - organized to match CALCULATIONS column grouping (rows 7-8 are MARKET RATING and CONSENSUS PRICE)
  let row = 9;
  
  // Define section colors - Professional color scheme with better contrast
  const SECTION_COLORS = {
    PRICE_VOLUME: '#2563EB',    // Blue - Primary data
    PERFORMANCE: '#7C3AED',     // Purple - Performance metrics
    TREND: '#059669',           // Green - Trend indicators
    MOMENTUM: '#DC2626',        // Red - Momentum signals
    VOLATILITY: '#EA580C',      // Orange - Volatility measures
    TARGET: '#DB2777'           // Pink - Target levels
  };
  
  // SIGNALING Section (B-D) - Already displayed in rows 4-6, skip to avoid duplication
  
  // PRICE / VOLUME Section (G-I) - CORRECTED COLUMN REFERENCES
  row = addSectionWithColor_(REPORT, row, 'PRICE / VOLUME', SECTION_COLORS.PRICE_VOLUME);
  row = addDataRow_(REPORT, row, 'PRICE', lookup('G'), '$#,##0.00');
  row = addDataRow_(REPORT, row, 'CHG%', lookup('H'), '0.00%');
  row = addDataRow_(REPORT, row, 'Vol Trend', lookup('I'), '0.00"x"');
  row = addDataRow_(REPORT, row, 'P/E', '=IFERROR(GOOGLEFINANCE($A$1,"pe"),"")', '0.00');
  row = addDataRow_(REPORT, row, 'EPS', '=IFERROR(GOOGLEFINANCE($A$1,"eps"),"")', '0.00');
  row = addDataRow_(REPORT, row, 'Range %', `=IFERROR(IF(AND(ISNUMBER(VALUE(LEFT(A2,LEN(A2)-1))),ISNUMBER(VALUE(LEFT(B2,LEN(B2)-1))),ISNUMBER(VALUE(LEFT(C2,LEN(C2)-1)))),LET(currentPrice,GOOGLEFINANCE($A$1,"price"),historicalDate,A3,historicalPrice,INDEX(GOOGLEFINANCE($A$1,"price",historicalDate),2,2),(currentPrice/historicalPrice-1)),"Select Date"),"—")`, '0.00%');
  
  // PERFORMANCE Section (J-M) - CORRECTED COLUMN REFERENCES
  row = addSectionWithColor_(REPORT, row, 'PERFORMANCE', SECTION_COLORS.PERFORMANCE);
  row = addDataRow_(REPORT, row, 'ATH (TRUE)', lookup('J'), '$#,##0.00');
  row = addDataRow_(REPORT, row, 'ATH Diff %', lookup('K'), '0.00%');
  row = addDataRow_(REPORT, row, 'ATH ZONE', lookup('L'), '@');
  row = addDataRow_(REPORT, row, 'FUNDAMENTAL', lookup('M'), '@');
  
  // TREND Section (N-Q) - CORRECTED COLUMN REFERENCES
  row = addSectionWithColor_(REPORT, row, 'TREND', SECTION_COLORS.TREND);
  row = addDataRow_(REPORT, row, 'TREND STATE', lookup('N'), '@');
  row = addDataRow_(REPORT, row, 'SMA 20', lookup('O'), '$#,##0.00');
  row = addDataRow_(REPORT, row, 'SMA 50', lookup('P'), '$#,##0.00');
  row = addDataRow_(REPORT, row, 'SMA 200', lookup('Q'), '$#,##0.00');
  
  // MOMENTUM Section (R-V) - CORRECTED COLUMN REFERENCES
  row = addSectionWithColor_(REPORT, row, 'MOMENTUM', SECTION_COLORS.MOMENTUM);
  row = addDataRow_(REPORT, row, 'RSI', lookup('R'), '0.0');
  row = addDataRow_(REPORT, row, 'MACD Hist', lookup('S'), '0.000');
  row = addDataRow_(REPORT, row, 'Divergence', lookup('T'), '@');
  row = addDataRow_(REPORT, row, 'ADX (14)', lookup('U'), '0.00');
  row = addDataRow_(REPORT, row, 'Stoch %K (14)', lookup('V'), '0.00%');
  
  // VOLATILITY Section (W-Z) - CORRECTED COLUMN REFERENCES
  row = addSectionWithColor_(REPORT, row, 'VOLATILITY', SECTION_COLORS.VOLATILITY);
  row = addDataRow_(REPORT, row, 'VOL REGIME', lookup('W'), '@');
  row = addDataRow_(REPORT, row, 'BBP SIGNAL', lookup('X'), '@');
  row = addDataRow_(REPORT, row, 'ATR (14)', lookup('Y'), '0.00');
  row = addDataRow_(REPORT, row, 'Bollinger %B', lookup('Z'), '0.0%');
  
  // TARGET Section (AA-AG) - CORRECTED COLUMN REFERENCES
  row = addSectionWithColor_(REPORT, row, 'TARGET', SECTION_COLORS.TARGET);
  row = addDataRow_(REPORT, row, 'Target (3:1)', lookup('AA'), '$#,##0.00');
  row = addDataRow_(REPORT, row, 'R:R Quality', lookup('AB'), '0.00"x"');
  row = addDataRow_(REPORT, row, 'Support', lookup('AC'), '$#,##0.00');
  row = addDataRow_(REPORT, row, 'Resistance', lookup('AD'), '$#,##0.00');
  row = addDataRow_(REPORT, row, 'ATR STOP', lookup('AE'), '$#,##0.00');
  row = addDataRow_(REPORT, row, 'ATR TARGET', lookup('AF'), '$#,##0.00');
  row = addDataRow_(REPORT, row, 'POSITION SIZE', lookup('AG'), '@')
  
  // INSTITUTIONAL ANALYSIS - A43:C60
  // Use custom function DASH_REPORT instead of long formula to avoid length limits
  // The DASH_REPORT function must be added to Apps Script project separately
  
  // Set custom function formula on A43 FIRST
  REPORT.getRange('A43').setFormula('=DASH_REPORT(A1)');
  
  // Now merge and apply formatting
  REPORT.getRange('A43:C60').merge()
    .setBackground(P.BG_ROW_A)
    .setFontColor('#FFFF00') // Yellow font
    .setFontWeight('normal')
    .setFontSize(10)
    .setFontFamily('Calibri')
    .setWrap(true)
    .setVerticalAlignment('top')
    .setHorizontalAlignment('left')
    .setBorder(true, true, true, true, false, false, '#FFFFFF', SpreadsheetApp.BorderStyle.SOLID);
  
  // Add margin: 3 columns × 4 rows below the report (starting at row 61)
  const marginStartRow = 61;
  REPORT.getRange(marginStartRow, 1, 4, 3)
    .setBackground('#000000')
    .clearContent();
  
  // DON'T DELETE COLUMN D - keep all columns as they are
  // Column D is now used for chart controls and chart area
  
  // Set white borders for columns A, B, C (all data rows)
  const lastDataRow = Math.max(60, REPORT.getLastRow());
  REPORT.getRange(1, 1, lastDataRow, 3)
    .setBorder(true, true, true, true, true, true, '#FFFFFF', SpreadsheetApp.BorderStyle.SOLID);
  
  // Final styling - NO BORDERS for clean professional look
  REPORT.setHiddenGridlines(true);
  
  // Setup report layout (dark backgrounds for M1:N1, cell merges)
  // Called at the END to avoid merge conflicts with earlier cell operations
  setupReportLayout(REPORT);
  
  // Always create chart since PRICE checkbox is enabled by default
  createReportChart_(REPORT);
}

/**
 * Add section header
 */
function addSection_(REPORT, row, title) {
  const P = reportPalette___();
  REPORT.getRange(row, 1).setValue(title);
  REPORT.getRange(row, 1, 1, 3).merge()
    .setBackground('#1F2937') // Slightly lighter than PANEL for distinction
    .setFontColor(P.TEXT)
    .setFontWeight('normal')
    .setFontSize(10)
    .setFontFamily('Calibri')
    .setBorder(true, true, true, true, false, false, '#FFFFFF', SpreadsheetApp.BorderStyle.SOLID)
    .setHorizontalAlignment('left');
  REPORT.setRowHeight(row, 22);
  return row + 1;
}

/**
 * Add section header with custom color
 */
function addSectionWithColor_(REPORT, row, title, color) {
  const P = reportPalette___();
  REPORT.getRange(row, 1).setValue(title);
  REPORT.getRange(row, 1, 1, 3).merge()
    .setBackground(color)  // Use custom color instead of default
    .setFontColor('#FFFFFF')  // White text for better contrast
    .setFontWeight('normal')
    .setFontSize(10)
    .setFontFamily('Calibri')
    .setBorder(true, true, true, true, false, false, '#FFFFFF', SpreadsheetApp.BorderStyle.SOLID)
    .setHorizontalAlignment('left');
  REPORT.setRowHeight(row, 22);
  return row + 1;
}

/**
 * Add data row with enhanced formatting for rows 35-45 (two columns with inferences)
 */
function addDataRow_(REPORT, row, label, formula, format) {
  const P = reportPalette___();
  
  // Check if this is in the enhanced inference section (rows 34-44)
  const isInferenceSection = (row >= 34 && row <= 44);
  
  // Label
  REPORT.getRange(row, 1).setValue(label);
  
  // Formula in column B
  REPORT.getRange(row, 2).setFormula(formula);
  
  // Enhanced two-column format for inference section (rows 34-44)
  if (isInferenceSection) {
    // Column C gets the inference/narrative
    const narrativeFormula = getNarrativeFormula_(label);
    REPORT.getRange(row, 3).setFormula(narrativeFormula);
    
    // Special styling for inference section
    const bg = P.BG_ROW_A; // Consistent background
    REPORT.getRange(row, 1, 1, 3).setBackground(bg);
    
    REPORT.getRange(row, 1).setFontColor(P.MUTED).setFontWeight('normal').setHorizontalAlignment('left').setFontFamily('Calibri');
    REPORT.getRange(row, 2).setFontColor(P.TEXT).setFontWeight('normal').setHorizontalAlignment('left').setFontFamily('Calibri');
    REPORT.getRange(row, 3).setFontColor(P.TEXT).setFontWeight('normal').setHorizontalAlignment('left').setWrap(true).setFontFamily('Calibri');
    
    REPORT.setRowHeight(row, 40); // Taller for inference content
  } else {
    // Original format for other rows - check if split zone (rows 7-44)
    const isSplit = (row >= 7 && row <= 44);
    if (isSplit) {
      const narrativeFormula = getNarrativeFormula_(label);
      REPORT.getRange(row, 3).setFormula(narrativeFormula);
      
      // NO MERGE - keep B and C separate for split zone
    } else {
      // Merge B:C for non-split rows (rows beyond 44)
      REPORT.getRange(row, 2, 1, 2).merge();
    }
    
    // Styling
    const bg = ((row % 2) === 0) ? P.BG_ROW_A : P.BG_ROW_B;
    REPORT.getRange(row, 1, 1, 3).setBackground(bg);
    
    REPORT.getRange(row, 1).setFontColor(P.MUTED).setFontWeight('normal').setHorizontalAlignment('left').setFontFamily('Calibri');
    REPORT.getRange(row, 2).setFontColor(P.TEXT).setFontWeight('normal').setHorizontalAlignment('left').setFontFamily('Calibri');
    
    if (isSplit) {
      REPORT.getRange(row, 3).setFontColor(P.TEXT).setFontWeight('normal').setHorizontalAlignment('left').setWrap(true).setFontFamily('Calibri');
      REPORT.setRowHeight(row, 34); // Taller for narrative
    } else {
      REPORT.getRange(row, 2, 1, 2).setFontColor(P.TEXT).setFontWeight('normal').setHorizontalAlignment('left').setFontFamily('Calibri');
      REPORT.setRowHeight(row, 18);
    }
  }
  
  // Number formatting
  if (format) {
    REPORT.getRange(row, 2).setNumberFormat(format);
  }
  
  // Apply conditional formatting
  applyConditionalFormatting_(REPORT, row, label);
  
  // Special handling for SMA color coding (post-process after formula is set)
  if (label === 'SMA 20' || label === 'SMA 50' || label === 'SMA 200') {
    applySMAColorCoding_(REPORT, row, label);
  }
  
  // Special handling for Support/Resistance color coding
  if (label === 'Support' || label === 'Resistance') {
    applySupportResistanceColorCoding_(REPORT, row, label);
  }
  
  // Borders - Clean minimal style (only bottom border for separation) - WHITE borders
  REPORT.getRange(row, 1, 1, 3)
    .setBorder(false, false, true, false, false, false, '#FFFFFF', SpreadsheetApp.BorderStyle.SOLID);
  
  return row + 1;
}

/**
 * Get narrative formula for column C
 */
function getNarrativeFormula_(label) {
  const lookup = (col) => `IFERROR(INDEX(CALCULATIONS!${col}:${col},MATCH(UPPER(TRIM($A$1)),ARRAYFORMULA(UPPER(TRIM(CALCULATIONS!A:A))),0)),0)`;
  
  // Numeric lookup for calculations
  const numLookup = (col) => `IFERROR(VALUE(INDEX(CALCULATIONS!${col}:${col},MATCH(UPPER(TRIM($A$1)),ARRAYFORMULA(UPPER(TRIM(CALCULATIONS!A:A))),0))),0)`;
  
  switch (label) {
    // SIGNALING Section
    case 'SIGNAL':
      return '=""';
    
    case 'PATTERNS':
      return '=""';
    
    case 'DECISION':
      return '=""';
    
    // PRICE Section
    case 'PRICE':
      return '=""';
    
    case 'CHG%':
      return '=""';
    
    case 'P/E':
      return '=""';
    
    case 'EPS':
      return '=IFERROR(IF(GOOGLEFINANCE($A$1,"eps")>=0.50,"profitable","unprofitable"),"—")';
    
    case 'Range %':
      return '=IFERROR("since " & TEXT(A3,"yyyy-mm-dd"),"—")';
    
    // VOLUME Section
    case 'Vol Trend':
      return '=IFERROR(IF(' + lookup('I') + '>=1.5,"strong participation.",IF(' + lookup('I') + '>=1,"average participation.","low participation (drift risk).")),"—")';
    
    case 'VOL REGIME':
      return '=IFERROR("volatility environment.","—")';
    
    // PERFORMANCE Section
    case 'ATH (TRUE)':
      return '=""';
    
    case 'ATH Diff %':
      return '=IFERROR(IF(' + lookup('K') + '>=-0.02,"at ATH zone",IF(' + lookup('K') + '>=-0.15,"pullback zone","correction territory")),"—")';
    
    case 'ATH ZONE':
      return '=""';
    
    // FUNDAMENTAL Section
    case 'FUNDAMENTAL':
      return '=IFERROR(' + lookup('M') + ' & " valuation assessment.","—")';
    
    // TREND Section
    case 'TREND STATE':
      return '=IFERROR(' + lookup('N') + ' & " market regime based on SMA200 position.","—")';
    
    case 'SMA 20':
      return '=IFERROR(TEXT((' + numLookup('O') + '-' + numLookup('G') + ')/' + numLookup('G') + ',"+0%;-0%") & IF(' + numLookup('G') + '>=' + numLookup('O') + '," below price - short-term bullish."," above price - short-term bearish."),"—")';
    
    case 'SMA 50':
      return '=IFERROR(TEXT((' + numLookup('P') + '-' + numLookup('G') + ')/' + numLookup('G') + ',"+0%;-0%") & IF(' + numLookup('G') + '>=' + numLookup('P') + '," below price - medium-term bullish."," above price - medium-term bearish."),"—")';
    
    case 'SMA 200':
      return '=IFERROR(TEXT((' + numLookup('Q') + '-' + numLookup('G') + ')/' + numLookup('G') + ',"+0%;-0%") & IF(' + numLookup('G') + '>=' + numLookup('Q') + '," below price - RISK-ON regime."," above price - RISK-OFF regime."),"—")';
    
    // MOMENTUM Section
    case 'RSI':
      return '=IFERROR(IF(' + lookup('R') + '>=70,"overbought zone.",IF(' + lookup('R') + '<=30,"oversold zone.",IF(' + lookup('R') + '>=55,"positive momentum.",IF(' + lookup('R') + '<=45,"weak momentum.","neutral range.")))),"—")';
    
    case 'MACD Hist':
      return '=IFERROR(IF(' + lookup('S') + '>0,"positive momentum impulse.",IF(' + lookup('S') + '<0,"negative momentum impulse.","flat momentum.")),"—")';
    
    case 'Divergence':
      return '=IFERROR(IF(' + lookup('T') + '="BULL DIV","Bullish divergence detected.",IF(' + lookup('T') + '="BEAR DIV","Bearish divergence detected.","No divergence detected.")),"—")';
    
    case 'ADX (14)':
      return '=IFERROR("Trend strength: " & IF(' + lookup('S') + '>=25," strong ",IF(' + lookup('S') + '>=20," developing ",IF(' + lookup('S') + '>=15," weak "," range-bound "))),"—")';
    
    case 'Stoch %K (14)':
      return '=IFERROR(IF(' + lookup('T') + '>=0.8,"overbought timing.",IF(' + lookup('T') + '<=0.2,"oversold timing.","neutral timing.")),"—")';
    
    case 'BBP SIGNAL':
      return '=IFERROR("Signal for mean reversion opportunities.","—")';
    
    // VOLATILITY Section
    case 'ATR (14)':
      return '=IFERROR(TEXT(' + lookup('W') + '/' + lookup('E') + ',"0.0%") & " of price - volatility measure.","—")';
    
    case 'Bollinger %B':
      return '=IFERROR(IF(' + lookup('X') + '>1," above upper band.",IF(' + lookup('X') + '>=0.8," upper band zone.",IF(' + lookup('X') + '<0," below lower band.",IF(' + lookup('X') + '<=0.2," lower band zone."," mid-band zone.")))),"—")';
    
    // TARGET Section
    case 'Target (3:1)':
      return '=IFERROR(TEXT((' + numLookup('Y') + '/' + numLookup('E') + '-1),"+0.00%;-0.00%") & " upside potential","—")';
    
    case 'R:R Quality':
      return '=IFERROR(IF(' + lookup('Z') + '>=3," elite asymmetry.",IF(' + lookup('Z') + '>=1.5," acceptable asymmetry."," poor asymmetry.")),"—")';
    
    // LEVELS Section
    case 'Support':
      return '=IFERROR(IF(' + numLookup('AA') + '<' + numLookup('E') + ',TEXT((' + numLookup('E') + '/' + numLookup('AA') + '-1),"+0.0%") & " support",TEXT((' + numLookup('AA') + '/' + numLookup('E') + '-1),"+0.0%") & " support"),"—")';
    
    case 'Resistance':
      return '=IFERROR(IF(' + numLookup('AB') + '>' + numLookup('E') + ',TEXT((' + numLookup('E') + '/' + numLookup('AB') + '-1),"0.0%") & " below Resistance",TEXT((' + numLookup('E') + '/' + numLookup('AB') + '-1),"+0.0%") & " above Resistance"),"—")';
    
    case 'ATR STOP':
      return '=IFERROR(TEXT(ABS((' + numLookup('E') + '/' + numLookup('AC') + '-1)),"+0.0%;-0.0%") & " risk from current price","—")';
    
    case 'ATR TARGET':
      return '=IFERROR(TEXT((' + numLookup('AD') + '/' + numLookup('E') + '-1),"+0.0%;-0.0%") & " reward potential","—")';
    
    case 'POSITION SIZE':
      return '=""';
    
    default:
      return '"No explanation available."';
  }
}

/**
 * Apply conditional formatting based on values
 */
function applyConditionalFormatting_(REPORT, row, label) {
  const P = reportPalette___();
  const valueCell = REPORT.getRange(row, 2);
  
  // Create conditional formatting rules based on label
  const rules = [];
  
  switch (label.toUpperCase()) {
    case 'CHG%':
      // Positive change = green, negative = red
      rules.push(
        SpreadsheetApp.newConditionalFormatRule()
          .whenNumberGreaterThan(0)
          .setBackground(P.CHIP_POS)
          .setFontColor(P.POS_TXT)
          .setRanges([valueCell])
          .build(),
        SpreadsheetApp.newConditionalFormatRule()
          .whenNumberLessThan(0)
          .setBackground(P.CHIP_NEG)
          .setFontColor(P.NEG_TXT)
          .setRanges([valueCell])
          .build()
      );
      break;
      
    case 'P/E':
      rules.push(
        SpreadsheetApp.newConditionalFormatRule()
          .whenNumberLessThanOrEqualTo(25)
          .setBackground(P.CHIP_POS)
          .setFontColor(P.POS_TXT)
          .setRanges([valueCell])
          .build(),
        SpreadsheetApp.newConditionalFormatRule()
          .whenNumberBetween(25.01, 35)
          .setBackground(P.CHIP_WARN)
          .setFontColor(P.WARN_TXT)
          .setRanges([valueCell])
          .build(),
        SpreadsheetApp.newConditionalFormatRule()
          .whenNumberGreaterThan(35)
          .setBackground(P.CHIP_NEG)
          .setFontColor(P.NEG_TXT)
          .setRanges([valueCell])
          .build()
      );
      break;
      
    case 'EPS':
      rules.push(
        SpreadsheetApp.newConditionalFormatRule()
          .whenNumberGreaterThanOrEqualTo(0.5)
          .setBackground(P.CHIP_POS)
          .setFontColor(P.POS_TXT)
          .setRanges([valueCell])
          .build(),
        SpreadsheetApp.newConditionalFormatRule()
          .whenNumberBetween(0.01, 0.49)
          .setBackground(P.CHIP_WARN)
          .setFontColor(P.WARN_TXT)
          .setRanges([valueCell])
          .build(),
        SpreadsheetApp.newConditionalFormatRule()
          .whenNumberLessThanOrEqualTo(0)
          .setBackground(P.CHIP_NEG)
          .setFontColor(P.NEG_TXT)
          .setRanges([valueCell])
          .build()
      );
      break;
      
    case 'ATH DIFF %':
      rules.push(
        SpreadsheetApp.newConditionalFormatRule()
          .whenNumberGreaterThanOrEqualTo(-0.02)
          .setBackground(P.CHIP_POS)
          .setFontColor(P.POS_TXT)
          .setRanges([valueCell])
          .build(),
        SpreadsheetApp.newConditionalFormatRule()
          .whenNumberBetween(-0.15, -0.02)
          .setBackground(P.CHIP_WARN)
          .setFontColor(P.WARN_TXT)
          .setRanges([valueCell])
          .build(),
        SpreadsheetApp.newConditionalFormatRule()
          .whenNumberLessThan(-0.15)
          .setBackground(P.CHIP_NEG)
          .setFontColor(P.NEG_TXT)
          .setRanges([valueCell])
          .build()
      );
      break;
      
    case 'RANGE %':
      rules.push(
        SpreadsheetApp.newConditionalFormatRule()
          .whenNumberGreaterThan(0)
          .setBackground(P.CHIP_POS)
          .setFontColor(P.POS_TXT)
          .setRanges([valueCell])
          .build(),
        SpreadsheetApp.newConditionalFormatRule()
          .whenNumberLessThan(0)
          .setBackground(P.CHIP_NEG)
          .setFontColor(P.NEG_TXT)
          .setRanges([valueCell])
          .build()
      );
      break;
      
    case 'VOL TREND':
      rules.push(
        SpreadsheetApp.newConditionalFormatRule()
          .whenNumberGreaterThanOrEqualTo(1.5)
          .setBackground(P.CHIP_POS)
          .setFontColor(P.POS_TXT)
          .setRanges([valueCell])
          .build(),
        SpreadsheetApp.newConditionalFormatRule()
          .whenNumberLessThan(1.0)
          .setBackground(P.CHIP_WARN)
          .setFontColor(P.WARN_TXT)
          .setRanges([valueCell])
          .build()
      );
      break;
      
    case 'R:R QUALITY':
      rules.push(
        SpreadsheetApp.newConditionalFormatRule()
          .whenNumberGreaterThanOrEqualTo(3)
          .setBackground(P.CHIP_POS)
          .setFontColor(P.POS_TXT)
          .setRanges([valueCell])
          .build(),
        SpreadsheetApp.newConditionalFormatRule()
          .whenNumberLessThan(1.5)
          .setBackground(P.CHIP_NEG)
          .setFontColor(P.NEG_TXT)
          .setRanges([valueCell])
          .build(),
        SpreadsheetApp.newConditionalFormatRule()
          .whenNumberBetween(1.5, 2.99)
          .setBackground(P.CHIP_WARN)
          .setFontColor(P.WARN_TXT)
          .setRanges([valueCell])
          .build()
      );
      break;
      
    case 'RSI':
      rules.push(
        SpreadsheetApp.newConditionalFormatRule()
          .whenNumberGreaterThanOrEqualTo(70)
          .setBackground(P.CHIP_WARN)
          .setFontColor(P.WARN_TXT)
          .setRanges([valueCell])
          .build(),
        SpreadsheetApp.newConditionalFormatRule()
          .whenNumberLessThanOrEqualTo(30)
          .setBackground(P.CHIP_POS)
          .setFontColor(P.POS_TXT)
          .setRanges([valueCell])
          .build(),
        SpreadsheetApp.newConditionalFormatRule()
          .whenNumberBetween(55, 69)
          .setBackground(P.CHIP_POS)
          .setFontColor(P.POS_TXT)
          .setRanges([valueCell])
          .build(),
        SpreadsheetApp.newConditionalFormatRule()
          .whenNumberBetween(31, 45)
          .setBackground(P.CHIP_WARN)
          .setFontColor(P.WARN_TXT)
          .setRanges([valueCell])
          .build()
      );
      break;
      
    case 'ADX (14)':
      rules.push(
        SpreadsheetApp.newConditionalFormatRule()
          .whenNumberGreaterThanOrEqualTo(25)
          .setBackground(P.CHIP_POS)
          .setFontColor(P.POS_TXT)
          .setRanges([valueCell])
          .build(),
        SpreadsheetApp.newConditionalFormatRule()
          .whenNumberBetween(20, 24)
          .setBackground(P.CHIP_WARN)
          .setFontColor(P.WARN_TXT)
          .setRanges([valueCell])
          .build()
      );
      break;
      
    case 'STOCH %K (14)':
      rules.push(
        SpreadsheetApp.newConditionalFormatRule()
          .whenNumberGreaterThanOrEqualTo(0.8)
          .setBackground(P.CHIP_WARN)
          .setFontColor(P.WARN_TXT)
          .setRanges([valueCell])
          .build(),
        SpreadsheetApp.newConditionalFormatRule()
          .whenNumberLessThanOrEqualTo(0.2)
          .setBackground(P.CHIP_POS)
          .setFontColor(P.POS_TXT)
          .setRanges([valueCell])
          .build()
      );
      break;
      
    case 'BOLLINGER %B':
      rules.push(
        SpreadsheetApp.newConditionalFormatRule()
          .whenNumberGreaterThan(1)
          .setBackground(P.CHIP_WARN)
          .setFontColor(P.WARN_TXT)
          .setRanges([valueCell])
          .build(),
        SpreadsheetApp.newConditionalFormatRule()
          .whenNumberGreaterThanOrEqualTo(0.8)
          .setBackground(P.CHIP_WARN)
          .setFontColor(P.WARN_TXT)
          .setRanges([valueCell])
          .build(),
        SpreadsheetApp.newConditionalFormatRule()
          .whenNumberLessThan(0)
          .setBackground(P.CHIP_POS)
          .setFontColor(P.POS_TXT)
          .setRanges([valueCell])
          .build(),
        SpreadsheetApp.newConditionalFormatRule()
          .whenNumberLessThanOrEqualTo(0.2)
          .setBackground(P.CHIP_POS)
          .setFontColor(P.POS_TXT)
          .setRanges([valueCell])
          .build()
      );
      break;
      
    case 'SMA 20':
    case 'SMA 50':  
    case 'SMA 200':
      // Note: For SMA comparison, we'd need cross-sheet references which aren't allowed
      // The color coding will be handled by the narrative text instead
      break;
      
    case 'MACD HIST':
      rules.push(
        SpreadsheetApp.newConditionalFormatRule()
          .whenNumberGreaterThan(0)
          .setBackground(P.CHIP_POS)
          .setFontColor(P.POS_TXT)
          .setRanges([valueCell])
          .build(),
        SpreadsheetApp.newConditionalFormatRule()
          .whenNumberLessThan(0)
          .setBackground(P.CHIP_NEG)
          .setFontColor(P.NEG_TXT)
          .setRanges([valueCell])
          .build()
      );
      break;
      
    case 'SUPPORT':
    case 'RESISTANCE':
      // Note: Support/Resistance comparison requires cross-sheet references which aren't allowed
      // The color coding will be handled by the narrative text instead
      break;
      
    case 'TREND STATE':
      rules.push(
        SpreadsheetApp.newConditionalFormatRule()
          .whenTextEqualTo('BULL')
          .setBackground(P.CHIP_POS)
          .setFontColor(P.POS_TXT)
          .setRanges([valueCell])
          .build(),
        SpreadsheetApp.newConditionalFormatRule()
          .whenTextEqualTo('BEAR')
          .setBackground(P.CHIP_NEG)
          .setFontColor(P.NEG_TXT)
          .setRanges([valueCell])
          .build()
      );
      break;
      
    case 'VOL REGIME':
      rules.push(
        SpreadsheetApp.newConditionalFormatRule()
          .whenTextEqualTo('LOW VOL')
          .setBackground(P.CHIP_POS)
          .setFontColor(P.POS_TXT)
          .setRanges([valueCell])
          .build(),
        SpreadsheetApp.newConditionalFormatRule()
          .whenTextEqualTo('EXTREME VOL')
          .setBackground(P.CHIP_NEG)
          .setFontColor(P.NEG_TXT)
          .setRanges([valueCell])
          .build(),
        SpreadsheetApp.newConditionalFormatRule()
          .whenTextEqualTo('HIGH VOL')
          .setBackground(P.CHIP_WARN)
          .setFontColor(P.WARN_TXT)
          .setRanges([valueCell])
          .build()
      );
      break;
      
    case 'ATH ZONE':
      rules.push(
        SpreadsheetApp.newConditionalFormatRule()
          .whenTextEqualTo('AT ATH')
          .setBackground(P.CHIP_POS)
          .setFontColor(P.POS_TXT)
          .setRanges([valueCell])
          .build(),
        SpreadsheetApp.newConditionalFormatRule()
          .whenTextEqualTo('DEEP VALUE ZONE')
          .setBackground(P.CHIP_POS)
          .setFontColor(P.POS_TXT)
          .setRanges([valueCell])
          .build(),
        SpreadsheetApp.newConditionalFormatRule()
          .whenTextContains('RESISTANCE')
          .setBackground(P.CHIP_WARN)
          .setFontColor(P.WARN_TXT)
          .setRanges([valueCell])
          .build()
      );
      break;
      
    case 'BBP SIGNAL':
      rules.push(
        SpreadsheetApp.newConditionalFormatRule()
          .whenTextContains('EXTREME OVERSOLD')
          .setBackground(P.CHIP_POS)
          .setFontColor(P.POS_TXT)
          .setRanges([valueCell])
          .build(),
        SpreadsheetApp.newConditionalFormatRule()
          .whenTextContains('EXTREME OVERBOUGHT')
          .setBackground(P.CHIP_NEG)
          .setFontColor(P.NEG_TXT)
          .setRanges([valueCell])
          .build(),
        SpreadsheetApp.newConditionalFormatRule()
          .whenTextContains('MOMENTUM STRONG')
          .setBackground(P.CHIP_POS)
          .setFontColor(P.POS_TXT)
          .setRanges([valueCell])
          .build()
      );
      break;
  }
  
  // Apply the rules
  if (rules.length > 0) {
    const sheet = valueCell.getSheet();
    const existingRules = sheet.getConditionalFormatRules();
    sheet.setConditionalFormatRules(existingRules.concat(rules));
  }
}

/**
 * Add narrative section
 */
function addNarrative_(REPORT, row, title, formula) {
  const P = reportPalette___();
  
  // Header
  REPORT.getRange(row, 1).setValue(title);
  REPORT.getRange(row, 1, 1, 3).merge()
    .setBackground(P.PANEL)
    .setFontColor(P.TEXT)
    .setFontWeight('normal')
    .setFontSize(11)
    .setFontFamily('Calibri')
    .setBorder(true, true, true, true, false, false, P.GRID, SpreadsheetApp.BorderStyle.SOLID)
    .setHorizontalAlignment('left');
  REPORT.setRowHeight(row, 22);
  row++;
  
  // Content
  REPORT.getRange(row, 1).setFormula(formula);
  REPORT.getRange(row, 1, 1, 3).merge()
    .setBackground('#0E1624')
    .setFontColor(P.WARN_TXT) // Changed to yellow
    .setVerticalAlignment('top')
    .setFontSize(10)
    .setFontFamily('Calibri')
    .setWrap(true)
    .setHorizontalAlignment('left')
    .setFontWeight('normal'); // Remove bold font
  REPORT.setRowHeight(row, 120);
  
  return row + 2;
}

/**
 * Setup ticker dropdown - now sets ticker name in merged A1:C1
 */
function setupReportTickerDropdown_(reportSheet, inputSheet) {
  // Get DASHBOARD sheet instead of using INPUT sheet
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dashboardSheet = ss.getSheetByName('DASHBOARD');
  
  if (!dashboardSheet) {
    console.log('DASHBOARD sheet not found, falling back to INPUT sheet');
    // Fallback to original INPUT sheet logic if DASHBOARD doesn't exist
    const last = inputSheet.getLastRow();
    const height = Math.max(1, last - 2);
    const rng = inputSheet.getRange(3, 1, height, 1);
    
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInRange(rng, true)
      .setAllowInvalid(true)
      .build();
    
    const a1 = reportSheet.getRange('A1');
    a1.setDataValidation(rule);a
    
    if (!a1.getValue()) {
      a1.setValue('AAPL');
    }
    
    SpreadsheetApp.flush();
    Utilities.sleep(100);
    return;
  }
  
  // Use DASHBOARD sheet A4:A range for ticker dropdown
  const dashboardLast = dashboardSheet.getLastRow();
  const dashboardHeight = Math.max(1, dashboardLast - 3); // Start from row 4, so subtract 3
  const dashboardRng = dashboardSheet.getRange(4, 1, dashboardHeight, 1); // A4:A range

  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(dashboardRng, true)
    .setAllowInvalid(true)
    .build();

  // Set validation on merged A1:C1 - use A1 as the reference cell
  const a1 = reportSheet.getRange('A1');
  a1.setDataValidation(rule);
  
  // Set default ticker to first ticker from DASHBOARD!A4 instead of hardcoded 'AAPL'
  if (!a1.getValue()) {
    const firstTicker = dashboardSheet.getRange('A4').getValue();
    if (firstTicker && String(firstTicker).trim()) {
      a1.setValue(String(firstTicker).trim());
    } else {
      a1.setValue('AAPL'); // Fallback if DASHBOARD A4 is empty
    }
  }
  
  SpreadsheetApp.flush();
  Utilities.sleep(100);
}

/**
 * Setup chart control checkboxes - all in row 1 starting from E1
 */
/**
 * Setup chart control checkboxes - UPDATED VERSION without date controls (now in A2:D2)
 */
function setupChartControls_(REPORT) {
  const P = reportPalette___();
  
  // Clear row 1 and 2 from D to M
  REPORT.getRange('D1:M2').clearContent().clearFormat();
  
  // All 9 controls in consecutive columns D through L
  const controls = [
    ['PRICE', true],  // Enable by default to create chart on first load
    ['SMA20', true],  // Enable by default
    ['SMA50', true],  // Enable by default
    ['SMA200', true], // Enable by default
    ['VOLUME', true],  // Enable volume by default to show bull/bear bars
    ['SUPPORT', true],
    ['RESISTANCE', true],
    ['ATR STOP', false],
    ['ATR TARGET', false]
  ];
  
  for (let i = 0; i < controls.length; i++) {
    const [label, defaultValue] = controls[i];
    const col = 4 + i; // D=4, E=5, F=6, G=7, H=8, I=9, J=10, K=11, L=12
    
    // Set column width to ensure all are visible
    REPORT.setColumnWidth(col, 80);
    
    // Add label in row 1 with professional styling
    REPORT.getRange(1, col).setValue(label);
    REPORT.getRange(1, col)
      .setBackground(P.CONTROL_LABEL)
      .setFontColor(P.WHITE)
      .setFontWeight('normal')
      .setFontSize(9)
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
    
    // Add checkbox in row 2 with professional styling
    REPORT.getRange(2, col).insertCheckboxes();
    REPORT.getRange(2, col).setValue(defaultValue);
    REPORT.getRange(2, col)
      .setBackground(P.CONTROL_BG)
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
  }
}

/**
 * Setup chart section placeholder - prepare for floating chart at D3:N17
 * Add AI analysis below chart in D18
 * CRITICAL: Never clear D18 or below to protect AI formula
 */
function setupChartSection_(REPORT) {
  const P = reportPalette___();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Clear ONLY the chart display area D3:N17 - NEVER touch D18 or below
  REPORT.getRange('D3:N17').clearContent().clearFormat();
  
  // Set professional background for chart display area - NO BORDERS for clean look
  REPORT.getRange('D3:N17')
    .setBackground('#0F1419');  // Deep black for chart area
  
  // Add INPUT column F reference in D18:N42 merged area
  // Set formula for INPUT column F reference
  // ONLY set if formula doesn't already exist
  const d18 = REPORT.getRange('D18');
  const existingFormula = d18.getFormula();
  
  const locale = (ss.getSpreadsheetLocale() || "").toLowerCase();
  const SEP = (/^(en|en_)/.test(locale)) ? "," : ";";
  
  if (!existingFormula || !existingFormula.includes('=IFERROR')) {
    REPORT.getRange('D18:N42').merge()
      .setFormula(`=IFERROR(INDEX(INPUT!F:F${SEP}MATCH(UPPER(TRIM($A$1))${SEP}ARRAYFORMULA(UPPER(TRIM(INPUT!A:A)))${SEP}0))${SEP}"—")`)
      .setFontColor('#FDE047')  // Bright yellow for text
      .setBackground('#1A1D29')  // Dark background
      .setWrap(true)
      .setVerticalAlignment('top')
      .setHorizontalAlignment('left')
      .setFontSize(10)
      .setFontFamily('Calibri')
      .setFontWeight('normal');
  }
}

/**
 * Create dynamic chart using REPORT sheet data - PROPER FLOATING CHART WITH ENHANCED ERROR HANDLING
 */
function createReportChart_(REPORT) {
  try {
    // Remove all existing charts first to prevent overlapping
    const existingCharts = REPORT.getCharts();
    existingCharts.forEach(chart => REPORT.removeChart(chart));
    
    const ticker = String(REPORT.getRange('A1').getValue() || '').trim();
    if (!ticker) {
      console.log('No ticker selected, skipping chart creation');
      return;
    }
    
    console.log(`Starting chart creation for ticker: ${ticker}`);
    
    // Call the main chart creation logic
    createReportChartInternal_(REPORT, ticker);
    
  } catch (e) {
    console.log(`Critical error in createReportChart_: ${e.toString()}`);
    
    // Try to show a simple error message in the chart area - ONLY clear D3:N17, NOT D18 or below
    try {
      REPORT.getRange('D3:N17').clearContent().clearFormat();
      REPORT.getRange('D3').setValue(`Chart Error: ${e.toString().substring(0, 100)}`);
      REPORT.getRange('D3:N17')
        .setBackground('#2A0B0B')
        .setFontColor('#F87171')
        .setFontSize(10)
        .setWrap(true)
        .setVerticalAlignment('middle')
        .setHorizontalAlignment('center');
    } catch (errorDisplayError) {
      console.log(`Could not display error message: ${errorDisplayError.toString()}`);
    }
  }
}

/**
 * Internal chart creation logic - separated for better error handling
 */
function createReportChartInternal_(REPORT, ticker) {
  
  // Get checkbox states from row 2, columns D-L (4-12) AFTER column D deletion
  // BEFORE deletion they were in E-M (5-13), but column D was deleted
  const checkboxes = {
    PRICE: REPORT.getRange(2, 4).getValue() || false,      // Was E2, now D2
    SMA20: REPORT.getRange(2, 5).getValue() || false,      // Was F2, now E2
    SMA50: REPORT.getRange(2, 6).getValue() || false,      // Was G2, now F2
    SMA200: REPORT.getRange(2, 7).getValue() || false,     // Was H2, now G2
    VOLUME: REPORT.getRange(2, 8).getValue() || false,     // Was I2, now H2
    SUPPORT: REPORT.getRange(2, 9).getValue() || false,    // Was J2, now I2
    RESISTANCE: REPORT.getRange(2, 10).getValue() || false, // Was K2, now J2
    ATR_STOP: REPORT.getRange(2, 11).getValue() || false,  // Was L2, now K2
    ATR_TARGET: REPORT.getRange(2, 12).getValue() || false // Was M2, now L2
  };
  
  // Debug checkbox states
  console.log(`Checkbox states: ${JSON.stringify(checkboxes)}`);
  
  // Get date and interval controls from A2:C3 (updated location)
  const yearsStr = String(REPORT.getRange('A2').getValue() || '0Y');
  const monthsStr = String(REPORT.getRange('B2').getValue() || '1M');
  const daysStr = String(REPORT.getRange('C2').getValue() || '0D');
  const interval = String(REPORT.getRange('C3').getValue() || 'Daily').toUpperCase(); // Moved to C3
  
  // Extract numeric values from Y/M/D format
  const years = Number(yearsStr.replace('Y', '')) || 0;
  const months = Number(monthsStr.replace('M', '')) || 1;
  const days = Number(daysStr.replace('D', '')) || 0;
  
  // Calculate start date based on dropdown selections
  const currentDate = new Date();
  const startDate = new Date(currentDate.getFullYear() - years, currentDate.getMonth() - months, currentDate.getDate() - days);
  const isWeekly = interval === 'WEEKLY';
  
  // Remove existing charts first
  REPORT.getCharts().forEach(chart => REPORT.removeChart(chart));
  
  // Get real data from DATA sheet (following updateDynamicChart pattern)
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getSheetByName("DATA");
  if (!dataSheet) {
    console.log('DATA sheet not found');
    return;
  }
  
  // Find ticker column in DATA sheet (same as updateDynamicChart)
  const dataHeaders = dataSheet.getRange(2, 1, 1, dataSheet.getLastColumn()).getValues()[0];
  const colIdx = dataHeaders.indexOf(ticker);
  if (colIdx === -1) {
    console.log(`Ticker ${ticker} not found in DATA sheet`);
    return;
  }
  
  // Pull 6 cols: date, open, high, low, close, volume (same as updateDynamicChart)
  const raw = dataSheet.getRange(1, colIdx + 1, dataSheet.getLastRow(), 6).getValues();
  
  // Debug: Check raw data structure
  console.log(`Raw data from DATA sheet: ${raw.length} rows`);
  if (raw.length > 5) {
    console.log(`Sample raw row 5: ${JSON.stringify(raw[5])}`);
    console.log(`Sample raw row 6: ${JSON.stringify(raw[6])}`);
  }
  
  // Get current values from CALCULATIONS sheet for support/resistance/ATR/ATR STOP/ATR TARGET
  const CALC = ss.getSheetByName('CALCULATIONS');
  let support = 0, resistance = 0, atr = 0, currentRSI = 50, currentPrice = 0, atrStop = 0, atrTarget = 0;
  
  if (CALC) {
    const calcData = CALC.getDataRange().getValues();
    const tickerRow = calcData.findIndex(row => String(row[0]).toUpperCase().trim() === ticker.toUpperCase());
    if (tickerRow !== -1) {
      const calcRow = calcData[tickerRow];
      // Column indices are 0-based: A=0, B=1, C=2, ..., Z=25, AA=26, AB=27, AC=28, AD=29, AE=30, AF=31
      currentPrice = Number(calcRow[4]) || 0; // Column E (index 4) - Price
      support = Number(calcRow[28]) || 0; // Column AC (index 28) - Support
      resistance = Number(calcRow[29]) || 0; // Column AD (index 29) - Resistance
      atr = Number(calcRow[24]) || 0; // Column Y (index 24) - ATR (14)
      currentRSI = Number(calcRow[17]) || 0; // Column R (index 17) - RSI
      atrStop = Number(calcRow[30]) || 0; // Column AE (index 30) - ATR STOP
      atrTarget = Number(calcRow[31]) || 0; // Column AF (index 31) - ATR TARGET
      
      // Debug log to verify data is being read correctly
      console.log(`Chart data for ${ticker}: Price=${currentPrice}, Support=${support}, Resistance=${resistance}, ATR=${atr}, RSI=${currentRSI}, ATR_STOP=${atrStop}, ATR_TARGET=${atrTarget}`);
    } else {
      console.log(`Ticker ${ticker} not found in CALCULATIONS sheet`);
    }
  } else {
    console.log('CALCULATIONS sheet not found');
  }
  
  // Get live price from Google Finance API for today's data - ENHANCED ERROR HANDLING
  let livePrice = 0;
  try {
    // Use a more robust approach for getting live price
    const tempSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const tempCell = tempSheet.getRange('Z1'); // Use a safe temporary cell
    tempCell.setFormula(`=GOOGLEFINANCE("${ticker}","price")`);
    SpreadsheetApp.flush();
    Utilities.sleep(500); // Give time for API call to complete
    
    const rawValue = tempCell.getValue();
    livePrice = (typeof rawValue === 'number' && isFinite(rawValue) && rawValue > 0) ? rawValue : (currentPrice || 0);
    tempCell.clear(); // Clean up temporary formula
    
    console.log(`Live price for ${ticker}: ${livePrice} (raw: ${rawValue})`);
  } catch (e) {
    console.log(`Failed to get live price for ${ticker}: ${e.toString()}`);
    livePrice = currentPrice || 0;
  }
  
  // Fallback to current price if live price is invalid
  if (!livePrice || !isFinite(livePrice) || livePrice <= 0) {
    livePrice = currentPrice || 100; // Use reasonable fallback
    console.log(`Using fallback price: ${livePrice}`);
  }
  
  const sampleData = [];
  
  // Process historical data (following updateDynamicChart pattern) - ENHANCED ERROR HANDLING
  for (let i = 4; i < raw.length; i++) {
    const d = raw[i][0];
    const close = Number(raw[i][4]);
    const vol = Number(raw[i][5]) || 1000000; // Fallback volume if missing
    
    // Enhanced data validation
    if (!d || !(d instanceof Date) || !isFinite(close) || close < 0.01) {
      console.log(`Skipping invalid data row ${i}: date=${d}, close=${close}`);
      continue;
    }
    if (d < startDate) continue;
    if (isWeekly && d.getDay() !== 5) continue;
    
    // Debug volume data for first few rows
    if (i < 7) {
      console.log(`Processing row ${i}: date=${d}, close=${close}, volume=${vol}, prevClose=${(i > 4) ? Number(raw[i - 1][4]) : close}`);
    }
    
    // SMA Calculations (same as updateDynamicChart) - ENHANCED ERROR HANDLING
    try {
      const slice = raw.slice(Math.max(4, i - 200), i + 1).map(r => Number(r[4])).filter(n => isFinite(n) && n > 0);
      const s20 = slice.length >= 20 ? Number((slice.slice(-20).reduce((a, b) => a + b, 0) / 20).toFixed(2)) : null;
      const s50 = slice.length >= 50 ? Number((slice.slice(-50).reduce((a, b) => a + b, 0) / 50).toFixed(2)) : null;
      const s200 = slice.length >= 200 ? Number((slice.slice(-200).reduce((a, b) => a + b, 0) / 200).toFixed(2)) : null;
      
      // Use ATR STOP and ATR TARGET values directly from CALCULATIONS sheet (columns AE and AF)
      // These are static values (straight lines) calculated once for the ticker
      // No need to recalculate - just use the pre-calculated values
      
      // Build data row based on selected checkboxes - FIXED ORDER TO PREVENT SHIFTING
      const dataRow = [d];
      
      // Always add price first if selected
      if (checkboxes.PRICE) dataRow.push(close);
      
      // Add SMAs next - ALWAYS add if checkbox is checked to maintain column alignment
      if (checkboxes.SMA20) dataRow.push(s20 || 0);
      if (checkboxes.SMA50) dataRow.push(s50 || 0);
      if (checkboxes.SMA200) dataRow.push(s200 || 0);
      
      // Add support/resistance levels - ALWAYS add if checkbox is checked to maintain column alignment
      if (checkboxes.RESISTANCE) {
        dataRow.push(resistance || 0);
        if (i < 7) console.log(`Row ${i}: Adding resistance=${resistance || 0}`);
      }
      if (checkboxes.SUPPORT) {
        dataRow.push(support || 0);
        if (i < 7) console.log(`Row ${i}: Adding support=${support || 0}`);
      }
      
      // Add ATR levels - ALWAYS add if checkbox is checked to maintain column alignment
      if (checkboxes.ATR_STOP) dataRow.push(atrStop || 0);
      if (checkboxes.ATR_TARGET) dataRow.push(atrTarget || 0);
      
      // Add volume LAST to prevent series shifting issues - CORRECTED LOGIC
      if (checkboxes.VOLUME) {
        // Use open price for intraday comparison (close vs open)
        // Fallback to previous close if open is unavailable
        let open = Number(raw[i][1]);
        if (!open || isNaN(open)) {
          open = (i > 4) ? Number(raw[i - 1][4]) : close;
          if (i < 7) {
            console.log(`Row ${i}: Warning - Open price unavailable, using fallback (${open})`);
          }
        }
        
        // Split volume into bull/bear based on intraday price movement
        // Bull volume (green) when close >= open
        // Bear volume (red) when close < open
        const isBullVolume = close >= open;
        
        // Validate volume data
        if (!vol || isNaN(vol) || vol < 0) {
          console.log(`Row ${i}: Warning - Invalid volume data (${vol}), setting to 0`);
          vol = 0;
        }
        
        // Use original volume values - secondary axis will handle scaling
        const bullVol = isBullVolume ? vol : 0;
        const bearVol = isBullVolume ? 0 : vol;
        
        // Debug volume calculation for first few rows
        if (i < 7) {
          console.log(`Row ${i}: open=${open}, close=${close}, vol=${vol}, isBull=${isBullVolume}, bullVol=${bullVol}, bearVol=${bearVol}`);
        }
        
        // Always add both bull and bear volume to maintain consistent column count
        dataRow.push(bullVol, bearVol);
      }
      
      sampleData.push(dataRow);
    } catch (e) {
      console.log(`Error processing data row ${i}: ${e.toString()}`);
      // Continue processing other rows
    }
  }
  
  // 5. LIVE-STITCH: Add Today's Data point if missing (following updateDynamicChart pattern) - ENHANCED ERROR HANDLING
  const today = new Date();
  const lastDateInSample = sampleData.length > 0 ? sampleData[sampleData.length - 1][0] : null;
  
  if (livePrice > 0 && (!lastDateInSample || lastDateInSample.toDateString() !== today.toDateString())) {
    console.log(`Adding live data point for ${ticker}: price=${livePrice}, date=${today.toDateString()}`);
    
    try {
      const lastHistClose = sampleData.length > 0 ? sampleData[sampleData.length - 1][1] : livePrice;
      
      // For live SMAs, calculate using historical data + current price (same as updateDynamicChart)
      const fullCloses = raw.slice(4).map(r => Number(r[4])).filter(n => isFinite(n) && n > 0);
      fullCloses.push(livePrice);
      
      const liveS20 = fullCloses.length >= 20 ? Number((fullCloses.slice(-20).reduce((a, b) => a + b, 0) / 20).toFixed(2)) : null;
      const liveS50 = fullCloses.length >= 50 ? Number((fullCloses.slice(-50).reduce((a, b) => a + b, 0) / 50).toFixed(2)) : null;
      const liveS200 = fullCloses.length >= 200 ? Number((fullCloses.slice(-200).reduce((a, b) => a + b, 0) / 200).toFixed(2)) : null;
      
      // Use ATR STOP and ATR TARGET values directly from CALCULATIONS sheet (columns AE and AF)
      // These are static values (straight lines) - already retrieved above from CALCULATIONS
      // No need to recalculate for live data - use the same pre-calculated values
      
      // Build live data row based on selected checkboxes - SAME ORDER AS HISTORICAL DATA
      const liveDataRow = [today];
      
      // Always add price first if selected
      if (checkboxes.PRICE) liveDataRow.push(livePrice);
      
      // Add SMAs next - ALWAYS add if checkbox is checked to maintain column alignment
      if (checkboxes.SMA20) liveDataRow.push(liveS20 || 0);
      if (checkboxes.SMA50) liveDataRow.push(liveS50 || 0);
      if (checkboxes.SMA200) liveDataRow.push(liveS200 || 0);
      
      // Add support/resistance levels - ALWAYS add if checkbox is checked to maintain column alignment
      if (checkboxes.RESISTANCE) liveDataRow.push(resistance || 0);
      if (checkboxes.SUPPORT) liveDataRow.push(support || 0);
      
      // Add ATR levels - ALWAYS add if checkbox is checked to maintain column alignment
      if (checkboxes.ATR_STOP) liveDataRow.push(atrStop || 0);
      if (checkboxes.ATR_TARGET) liveDataRow.push(atrTarget || 0);
      
      // Add volume LAST - FIXED: Check if non-trading day, then use GOOGLEFINANCE volume
      if (checkboxes.VOLUME) {
        let bullVol = 0;
        let bearVol = 0;
        
        // Check if today is a non-trading day by comparing current price to last historical close
        // If price hasn't changed, it's a non-trading day (market closed)
        const priceChange = Math.abs(livePrice - lastHistClose);
        const isNonTradingDay = priceChange < 0.01; // Price change less than 1 cent = non-trading day
        
        if (isNonTradingDay) {
          // Non-trading day - use volume = 0
          console.log(`Live volume: Non-trading day detected (price unchanged: ${livePrice} vs ${lastHistClose}), using volume=0`);
        } else {
          // Trading day - get actual volume from GOOGLEFINANCE
          try {
            const tempSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
            const tempCell = tempSheet.getRange('Z2'); // Use a safe temporary cell
            tempCell.setFormula(`=GOOGLEFINANCE("${ticker}","volume")`);
            SpreadsheetApp.flush();
            Utilities.sleep(500); // Give time for API call to complete
            
            const todayVolume = Number(tempCell.getValue()) || 0;
            tempCell.clear(); // Clean up temporary formula
            
            console.log(`Today's volume from GOOGLEFINANCE: ${todayVolume}`);
            
            if (todayVolume > 0) {
              // Market is open or has traded today - use actual volume
              // For live data, the open price is the last historical close (where today's trading started)
              const liveOpen = lastHistClose;
              const isBullVolume = livePrice >= liveOpen;
              
              bullVol = isBullVolume ? todayVolume : 0;
              bearVol = isBullVolume ? 0 : todayVolume;
              
              console.log(`Live volume: open=${liveOpen}, close=${livePrice}, volume=${todayVolume}, isBull=${isBullVolume}, bullVol=${bullVol}, bearVol=${bearVol}`);
            } else {
              console.log(`Live volume: GOOGLEFINANCE returned 0 volume`);
            }
          } catch (e) {
            console.log(`Error fetching today's volume: ${e.toString()}, using volume=0`);
            // On error, use 0 volume
          }
        }
        
        // Always add both bull and bear volume to maintain consistent column count
        liveDataRow.push(bullVol, bearVol);
      }
      
      sampleData.push(liveDataRow);
      console.log(`Added live data row: ${JSON.stringify(liveDataRow)}`);
    } catch (e) {
      console.log(`Error adding live data point: ${e.toString()}`);
      // Continue without live data if there's an error
    }
  }
  
  // Ensure we have at least some data to work with
  if (sampleData.length === 0) {
    console.log('No sample data available for chart creation');
    return;
  }
  
  // Build headers based on selected checkboxes - FIXED ORDER TO MATCH DATA
  const headers = ['Date'];
  if (checkboxes.PRICE) headers.push('Price');
  if (checkboxes.SMA20) headers.push('SMA20');
  if (checkboxes.SMA50) headers.push('SMA50');
  if (checkboxes.SMA200) headers.push('SMA200');
  if (checkboxes.RESISTANCE) headers.push('Resistance');
  if (checkboxes.SUPPORT) headers.push('Support');
  if (checkboxes.ATR_STOP) headers.push('ATR Stop');
  if (checkboxes.ATR_TARGET) headers.push('ATR Target');
  // Volume headers - always add both bull and bear when volume is enabled
  if (checkboxes.VOLUME) headers.push('Bull Volume', 'Bear Volume');
  
  console.log(`Fetched ${sampleData.length} real data points for ${ticker} from DATA sheet`);
  console.log(`Headers: ${headers.join(', ')}`);
  
  // Debug volume data if volume is enabled
  if (checkboxes.VOLUME && sampleData.length > 0) {
    const sampleRow = sampleData[Math.floor(sampleData.length / 2)];
    if (sampleRow.length >= 2) {
      const bullVol = sampleRow[sampleRow.length - 2];
      const bearVol = sampleRow[sampleRow.length - 1];
      console.log(`Sample volume data - Bull: ${bullVol}, Bear: ${bearVol}`);
    }
  }
  
  // Write chart data to hidden area (columns AA onwards) - ENHANCED ERROR HANDLING
  const startRow = 30;
  const startCol = 27; // Column AA
  
  try {
    // Clear and write data
    REPORT.getRange(startRow, startCol, 200, headers.length).clearContent();
    REPORT.getRange(startRow, startCol, 1, headers.length).setValues([headers]);
    REPORT.getRange(startRow + 1, startCol, sampleData.length, headers.length).setValues(sampleData);
    
    // Apply proper formatting to different column types
    // Date column (first column) - format as date
    REPORT.getRange(startRow + 1, startCol, sampleData.length, 1).setNumberFormat('dd/mm/yy');
    
    // Price/numeric columns (all except first column) - format as numbers
    if (headers.length > 1) {
      // Apply different formats for price vs volume columns
      let priceColCount = headers.length - 1;
      if (checkboxes.VOLUME) {
        priceColCount -= 2; // Subtract the 2 volume columns
        
        // Format price columns
        if (priceColCount > 0) {
          REPORT.getRange(startRow + 1, startCol + 1, sampleData.length, priceColCount).setNumberFormat('#,##0.00');
        }
        
        // Format volume columns with no decimals
        REPORT.getRange(startRow + 1, startCol + 1 + priceColCount, sampleData.length, 2).setNumberFormat('#,##0');
      } else {
        // All non-date columns are price columns
        REPORT.getRange(startRow + 1, startCol + 1, sampleData.length, priceColCount).setNumberFormat('#,##0.00');
      }
    }
  } catch (e) {
    console.log(`Error writing chart data: ${e.toString()}`);
    return; // Exit if we can't write data
  }
  
  // Debug: Log sample data to understand what's being written
  console.log(`Writing ${sampleData.length} rows with ${headers.length} columns to ${startRow},${startCol}`);
  console.log(`Headers: ${JSON.stringify(headers)}`);
  if (sampleData.length > 0) {
    console.log(`Sample row 1: ${JSON.stringify(sampleData[0])}`);
    if (sampleData.length > 1) {
      console.log(`Sample row 2: ${JSON.stringify(sampleData[1])}`);
    }
  }
  
  // Build series configuration based on checkboxes - FIXED ORDER TO PREVENT SHIFTING
  const seriesConfig = {};
  let seriesIndex = 0;
  
  // Add series in the same order as headers (excluding Date)
  if (checkboxes.PRICE) {
    seriesConfig[seriesIndex] = { 
      type: "line", 
      color: "#1A73E8", 
      lineWidth: 1, 
      labelInLegend: "Price",
      targetAxisIndex: 0 // Explicitly assign to primary axis
    };
    seriesIndex++;
  }
  
  if (checkboxes.SMA20) {
    seriesConfig[seriesIndex] = { 
      type: "line", 
      color: "#4CAF50", 
      lineWidth: 1.5, 
      labelInLegend: "SMA 20",
      targetAxisIndex: 0 // Explicitly assign to primary axis
    };
    seriesIndex++;
  }
  
  if (checkboxes.SMA50) {
    seriesConfig[seriesIndex] = { 
      type: "line", 
      color: "#9C27B0", 
      lineWidth: 1.5, 
      labelInLegend: "SMA 50",
      targetAxisIndex: 0 // Explicitly assign to primary axis
    };
    seriesIndex++;
  }
  
  if (checkboxes.SMA200) {
    seriesConfig[seriesIndex] = { 
      type: "line", 
      color: "#FF9800", 
      lineWidth: 2, 
      labelInLegend: "SMA 200",
      targetAxisIndex: 0 // Explicitly assign to primary axis
    };
    seriesIndex++;
  }
  
  if (checkboxes.RESISTANCE) {
    seriesConfig[seriesIndex] = { 
      type: "line", 
      color: "#34D399", // GREEN (not red) per requirements
      lineWidth: 2, // SOLID line (no lineDashStyle)
      labelInLegend: "Resistance",
      targetAxisIndex: 0 // Explicitly assign to primary axis
    };
    seriesIndex++;
  }
  
  if (checkboxes.SUPPORT) {
    seriesConfig[seriesIndex] = { 
      type: "line", 
      color: "#F87171", // RED (not green) per requirements
      lineWidth: 2, // SOLID line (no lineDashStyle)
      labelInLegend: "Support",
      targetAxisIndex: 0 // Explicitly assign to primary axis
    };
    seriesIndex++;
  }
  
  if (checkboxes.ATR_STOP) {
    seriesConfig[seriesIndex] = { 
      type: "line", 
      color: "#f80ed1ff",
      lineWidth: 1, 
      pointSize: 2, // This creates a "beaded" look
      lineDashStyle: 'long-dash', // DOTTED line per requirements - NOTE: Google Sheets may not render this in combo charts
      labelInLegend: "ATR Stop",
      targetAxisIndex: 0 // Explicitly assign to primary axis
    };
    seriesIndex++;
  }
  
  if (checkboxes.ATR_TARGET) {
    seriesConfig[seriesIndex] = { 
      type: "line", 
      color: "#e7ef0bff", 
      lineWidth: 1, 
      pointSize: 2, // This creates a "beaded" look
      lineDashStyle: 'long-dash', // DOTTED line per requirements - NOTE: Google Sheets may not render this in combo charts
      labelInLegend: "ATR Target",
      targetAxisIndex: 0 // Explicitly assign to primary axis
    };
    seriesIndex++;
  }
  
  // Add volume LAST - USE SECONDARY AXIS for proper visibility
  if (checkboxes.VOLUME) {
    // Bull volume series (green) - assign to secondary axis with proper opacity
    seriesConfig[seriesIndex] = { 
      type: "bars", 
      color: "#1976D2", // Medium Blue (not green)
      opacity: 0.6, // Reduced opacity for better visibility
      labelInLegend: "Bull Volume",
      targetAxisIndex: 1 // Secondary axis for volume
    };
    seriesIndex++;
    
    // Bear volume series (red) - assign to secondary axis with proper opacity
    seriesConfig[seriesIndex] = { 
      type: "bars", 
      color: "#C62828", // Darker red like updateDynamicChart()
      opacity: 0.6, // Reduced opacity for better visibility
      labelInLegend: "Bear Volume",
      targetAxisIndex: 1 // Secondary axis for volume
    };
    seriesIndex++;
  }
  
  // If no checkboxes are selected, show at least price as fallback
  if (Object.keys(seriesConfig).length === 0) {
    seriesConfig[0] = { type: "line", color: "#1A73E8", lineWidth: 1, labelInLegend: "Price" };
  }
  
  console.log(`Series config keys: ${Object.keys(seriesConfig).join(', ')}`);
  
  // Calculate scaling - CRITICAL FIX for volume visibility
  const allPriceValues = [];
  const allVolumeValues = [];
  
  sampleData.forEach(row => {
    for (let i = 1; i < row.length; i++) {
      if (typeof row[i] === 'number' && isFinite(row[i]) && row[i] > 0) {
        // Separate volume values from price values for proper scaling
        if (checkboxes.VOLUME && i >= row.length - 2) {
          allVolumeValues.push(row[i]);
        } else {
          allPriceValues.push(row[i]);
        }
      }
    }
  });
  
  // Include live price in price range calculation
  if (livePrice > 0) {
    allPriceValues.push(livePrice);
  }
  
  // Include support/resistance in price range if they exist
  if (support > 0) allPriceValues.push(support);
  if (resistance > 0) allPriceValues.push(resistance);
  
  const minPrice = allPriceValues.length > 0 ? Math.min(...allPriceValues) * 0.95 : 90;
  const maxPrice = allPriceValues.length > 0 ? Math.max(...allPriceValues) * 1.05 : 110;
  const priceRange = maxPrice - minPrice;
  
  // CRITICAL: Calculate proper volume scaling like updateDynamicChart()
  let maxVol = 1;
  if (checkboxes.VOLUME && allVolumeValues.length > 0) {
    maxVol = Math.max(...allVolumeValues.filter(v => isFinite(v)), 1);
    const minVol = Math.min(...allVolumeValues.filter(v => isFinite(v)), 0);
    
    console.log(`Volume scaling - max: ${maxVol}, min: ${minVol}`);
  }
  
  console.log(`Price range: ${minPrice} to ${maxPrice}`);
  
  // Setup dual axes - Primary for price, Secondary for volume with PROPER SCALING
  const vAxes = {
    0: { 
      viewWindow: { min: minPrice, max: maxPrice },
      format: '$#,##0.00',
      title: 'Price ($)',
      titleTextStyle: { color: '#FFFFFF', fontSize: 10 },
      textStyle: { color: '#FFFFFF', fontSize: 9 },
      gridlines: { color: '#374151' }
    },
    1: {
      viewWindow: { min: 0, max: maxVol * 4 }, // CRITICAL: Limit volume height like updateDynamicChart()
      format: '#,##0',
      title: 'Volume',
      titleTextStyle: { color: '#FFFFFF', fontSize: 10 },
      textStyle: { color: '#FFFFFF', fontSize: 9 },
      gridlines: { count: 0 } // Hide volume gridlines to avoid clutter
    }
  };
  
  // Primary vAxis configuration to ensure price format
  const primaryVAxis = {
    viewWindow: { min: minPrice, max: maxPrice },
    format: '$#,##0.00',
    title: 'Price ($)',
    titleTextStyle: { color: '#FFFFFF', fontSize: 10 },
    textStyle: { color: '#FFFFFF', fontSize: 9 },
    gridlines: { color: '#374151' }
  };
  
  console.log(`Price range: ${minPrice} to ${maxPrice}`);
  console.log(`VAxes config:`, JSON.stringify(vAxes));
  
  // Create floating chart exactly like updateDynamicChart - ENHANCED ERROR HANDLING
  // CRITICAL: Chart positioned at E3 with height 350px to fit E3:M17 only, NOT overwriting E18:M64 (AI analysis)
  try {
    let chart = REPORT.newChart()
      .setChartType(Charts.ChartType.COMBO)
      .addRange(REPORT.getRange(startRow, startCol, sampleData.length + 1, headers.length))
      .setOption("useFirstRowAsHeaders", true)
      .setOption("series", seriesConfig)
      .setOption("vAxes", vAxes)
      .setOption("vAxis", primaryVAxis) // Use explicit primary axis config
      .setOption("legend", { 
        position: "top",
        textStyle: { color: '#FFFFFF', fontSize: 10 }
      })
      .setOption("backgroundColor", '#0F172A')
      .setOption("hAxis", {
        textStyle: { color: '#FFFFFF', fontSize: 9 },
        gridlines: { color: '#374151' },
        format: 'dd/MM/yy'
      })
      .setPosition(3, 4, 0, 0) // Row 3, Column D
      .setOption("width", 720)  // Width to fit D3:N17
      .setOption("height", 350) // Height to fit D3:N17 only, NOT overwriting D18:N42
      .build();
    
    // Position and resize chart to span D3:N17 (was E3:O17 before column D deletion)
    chart = positionReportChart(chart);
    
    // Insert the floating chart
    REPORT.insertChart(chart);
    
    console.log(`Floating chart created for ${ticker} with ${Object.keys(seriesConfig).length} series, ${sampleData.length} data points, interval: ${interval}`);
  } catch (e) {
    console.log(`Error creating chart: ${e.toString()}`);
    // Try to create a simpler fallback chart with just price data
    try {
      console.log('Attempting to create fallback chart with price data only...');
      
      // Create minimal series config with just price
      const fallbackSeriesConfig = {
        0: { 
          type: "line", 
          color: "#1A73E8", 
          lineWidth: 1, 
          labelInLegend: "Price"
        }
      };
      
      const fallbackChart = REPORT.newChart()
        .setChartType(Charts.ChartType.LINE)
        .addRange(REPORT.getRange(startRow, startCol, sampleData.length + 1, Math.min(2, headers.length))) // Just date and price
        .setOption("useFirstRowAsHeaders", true)
        .setOption("series", fallbackSeriesConfig)
        .setOption("legend", { 
          position: "top",
          textStyle: { color: '#FFFFFF', fontSize: 10 }
        })
        .setOption("backgroundColor", '#0F172A')
        .setOption("hAxis", {
          textStyle: { color: '#FFFFFF', fontSize: 9 },
          gridlines: { color: '#374151' },
          format: 'dd/MM/yy'
        })
        .setOption("vAxis", {
          textStyle: { color: '#FFFFFF', fontSize: 9 },
          gridlines: { color: '#374151' },
          format: '$#,##0.00'
        })
        .setPosition(3, 4, 0, 0)  // Row 3, Column D
        .setOption("width", 720)
        .setOption("height", 350) // Height to fit D3:N17 only
        .build();
      
      REPORT.insertChart(fallbackChart);
      console.log(`Fallback chart created successfully for ${ticker}`);
    } catch (fallbackError) {
      console.log(`Fallback chart creation also failed: ${fallbackError.toString()}`);
      // If even the fallback fails, just log and continue
    }
  }
}

/**
 * Style chart series based on enabled checkboxes
 * Applies colors and line styles to ATR STOP, ATR TARGET, SUPPORT, and RESISTANCE series
 */
function styleChartSeries(chart, checkboxes) {
  // Get chart builder from existing chart
  const chartBuilder = chart.modify();
  
  // Track series index based on enabled checkboxes
  let seriesIndex = 0;
  
  // Price series (if enabled)
  if (checkboxes.PRICE) {
    seriesIndex++; // Price is series 0
  }
  
  // SMA series (if enabled)
  if (checkboxes.SMA20) seriesIndex++;
  if (checkboxes.SMA50) seriesIndex++;
  if (checkboxes.SMA200) seriesIndex++;
  
  // RESISTANCE series - GREEN color with SOLID line
  if (checkboxes.RESISTANCE) {
    chartBuilder.setOption(`series.${seriesIndex}.color`, '#34D399'); // Green
    chartBuilder.setOption(`series.${seriesIndex}.lineWidth`, 2);
    // Ensure solid line style (no dash style)
    seriesIndex++;
  }
  
  // SUPPORT series - RED color with SOLID line
  if (checkboxes.SUPPORT) {
    chartBuilder.setOption(`series.${seriesIndex}.color`, '#F87171'); // Red
    chartBuilder.setOption(`series.${seriesIndex}.lineWidth`, 2);
    // Ensure solid line style (no dash style)
    seriesIndex++;
  }
  
  // ATR STOP series - RED color with DOTTED line
  if (checkboxes.ATR_STOP) {
    chartBuilder.setOption(`series.${seriesIndex}.color`, '#F87171'); // Red
    chartBuilder.setOption(`series.${seriesIndex}.lineDashStyle`, [4, 4]); // Dotted
    chartBuilder.setOption(`series.${seriesIndex}.lineWidth`, 2);
    seriesIndex++;
  }
  
  // ATR TARGET series - GREEN color with DOTTED line
  if (checkboxes.ATR_TARGET) {
    chartBuilder.setOption(`series.${seriesIndex}.color`, '#34D399'); // Green
    chartBuilder.setOption(`series.${seriesIndex}.lineDashStyle`, [4, 4]); // Dotted
    chartBuilder.setOption(`series.${seriesIndex}.lineWidth`, 2);
    seriesIndex++;
  }
  
  return chartBuilder.build();
}

/**
 * Handle REPORT sheet edit events - called from Code.js onEdit trigger
 * This centralizes all REPORT sheet logic in one place
 */
function handleReportSheetEdit(e) {
  const range = e.range;
  const sheet = range.getSheet();
  const a1 = range.getA1Notation();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const row = range.getRow();
  const col = range.getColumn();
  
  // Handle chart controls: checkbox changes (row 2, columns D-L: 4-12), ticker change (A1), or date/interval change (A2:C3)
  if ((row === 2 && col >= 4 && col <= 12) || a1 === "A1" || (row === 2 && col >= 1 && col <= 3) || a1 === "C3") {
    try {
      ss.toast("🔄 Updating REPORT Chart...", "WORKING", 2);
      updateReportChart();
    } catch (err) {
      ss.toast("REPORT Chart update error: " + err.toString(), "⚠️ FAIL", 6);
    }
  }
}

/**
 * Update REPORT sheet chart - always recreate to ensure proper scaling
 */
function updateReportChart() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const REPORT = ss.getSheetByName('REPORT');
  if (!REPORT) return;
  
  // Add small delay to ensure cell values are updated
  SpreadsheetApp.flush();
  Utilities.sleep(100);
  
  // Always recreate chart to ensure proper axis scaling (especially for volume)
  // This is necessary because min/max data points change with different tickers/date ranges
  createReportChart_(REPORT);
}

/**
 * Set column widths - updated for new layout without D column
 */
function setReportColumnWidthsAndWrap___(REPORT) {
  const pxPerChar = 8;
  
  // A (ticker/dropdown), B (values), C (narratives)
  const colA = Math.max(70, Math.round(8 * pxPerChar + 15)); // Ticker column
  const colB = Math.max(70, Math.round(8 * pxPerChar + 15)); // Values column
  const colC = Math.max(170, Math.round(27 * pxPerChar + 20)); // Narratives column (wider)
  
  // Set main columns
  REPORT.setColumnWidth(1, colA);
  REPORT.setColumnWidth(2, colB);
  REPORT.setColumnWidth(3, colC);
  
  // Chart control columns (E-N) - handled in setupChartControls_
  
  const lastRow = Math.max(1, REPORT.getLastRow());
  REPORT.getRange(1, 1, lastRow, 13).setWrap(true); // Extended to column M
}

/**
 * Color palette
 */
function reportPalette___() {
  return {
    // Main backgrounds - Dark professional theme
    BG_TOP: '#1A1D29',        // Darker header background
    PANEL:  '#242938',        // Panel background
    BG_ROW_A: '#1E2230',      // Alternating row A (darker)
    BG_ROW_B: '#242938',      // Alternating row B (lighter)
    
    // Borders and grid
    GRID: '#3A3F51',          // Subtle grid lines
    
    // Text colors
    TEXT: '#E8EAED',          // Primary text (light gray)
    MUTED: '#9AA0A6',         // Muted/secondary text
    
    // Positive/Negative/Warning indicators
    POS_TXT:  '#34D399',      // Green for positive
    NEG_TXT:  '#F87171',      // Red for negative
    WARN_TXT: '#FBBF24',      // Amber for warning
    
    // Chip/badge backgrounds
    CHIP_POS:  '#064E3B',     // Dark green background
    CHIP_NEG:  '#7F1D1D',     // Dark red background
    CHIP_WARN: '#78350F',     // Dark amber background
    CHIP_NEU:  '#1E293B',     // Neutral dark blue
    
    // Special colors
    YELLOW: '#FDE047',        // Bright yellow for highlights
    BLACK:  '#0F1419',        // True black for contrast
    WHITE:  '#FFFFFF',        // Pure white
    
    // Chart control background
    CONTROL_BG: '#2D3748',    // Chart controls background
    CONTROL_LABEL: '#4A5568'  // Chart control labels
  };
}/**

 * Apply conditional formatting to decision cells (B4:B6) - updated for new layout with row 3 added
 */
function applyDecisionConditionalFormatting_(REPORT) {
  const P = reportPalette___();
  
  // SIGNAL cell (B4) - updated position
  const signalCell = REPORT.getRange('B4');
  const signalRules = [
    // Positive signals (Green)
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains('ATH BREAKOUT')
      .setBackground(P.CHIP_POS)
      .setFontColor(P.POS_TXT)
      .setRanges([signalCell])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains('VOLATILITY BREAKOUT')
      .setBackground(P.CHIP_POS)
      .setFontColor(P.POS_TXT)
      .setRanges([signalCell])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains('EXTREME OVERSOLD BUY')
      .setBackground(P.CHIP_POS)
      .setFontColor(P.POS_TXT)
      .setRanges([signalCell])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains('STRONG BUY')
      .setBackground(P.CHIP_POS)
      .setFontColor(P.POS_TXT)
      .setRanges([signalCell])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains('BUY')
      .setBackground(P.CHIP_POS)
      .setFontColor(P.POS_TXT)
      .setRanges([signalCell])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains('ACCUMULATE')
      .setBackground(P.CHIP_POS)
      .setFontColor(P.POS_TXT)
      .setRanges([signalCell])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains('BREAKOUT')
      .setBackground(P.CHIP_POS)
      .setFontColor(P.POS_TXT)
      .setRanges([signalCell])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains('MOMENTUM')
      .setBackground(P.CHIP_POS)
      .setFontColor(P.POS_TXT)
      .setRanges([signalCell])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains('UPTREND')
      .setBackground(P.CHIP_POS)
      .setFontColor(P.POS_TXT)
      .setRanges([signalCell])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains('BULLISH')
      .setBackground(P.CHIP_POS)
      .setFontColor(P.POS_TXT)
      .setRanges([signalCell])
      .build(),
    
    // Warning signals (Orange)
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains('VOLATILITY SQUEEZE')
      .setBackground(P.CHIP_WARN)
      .setFontColor(P.WARN_TXT)
      .setRanges([signalCell])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains('OVERSOLD')
      .setBackground(P.CHIP_WARN)
      .setFontColor(P.WARN_TXT)
      .setRanges([signalCell])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains('OVERBOUGHT')
      .setBackground(P.CHIP_WARN)
      .setFontColor(P.WARN_TXT)
      .setRanges([signalCell])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains('RANGE')
      .setBackground(P.CHIP_WARN)
      .setFontColor(P.WARN_TXT)
      .setRanges([signalCell])
      .build(),
    
    // Neutral signals (Gray)
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains('HOLD')
      .setBackground(P.CHIP_NEU)
      .setFontColor(P.TEXT)
      .setRanges([signalCell])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains('NEUTRAL')
      .setBackground(P.CHIP_NEU)
      .setFontColor(P.TEXT)
      .setRanges([signalCell])
      .build(),
    
    // Negative signals (Red)
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains('STOP OUT')
      .setBackground(P.CHIP_NEG)
      .setFontColor(P.NEG_TXT)
      .setRanges([signalCell])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains('RISK OFF')
      .setBackground(P.CHIP_NEG)
      .setFontColor(P.NEG_TXT)
      .setRanges([signalCell])
      .build()
  ];
  
  // FUNDAMENTAL cell (B5) - updated position
  const fundamentalCell = REPORT.getRange('B5');
  const fundamentalRules = [
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains('VALUE')
      .setBackground(P.CHIP_POS)
      .setFontColor(P.POS_TXT)
      .setRanges([fundamentalCell])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains('EXPENSIVE')
      .setBackground(P.CHIP_WARN)
      .setFontColor(P.WARN_TXT)
      .setRanges([fundamentalCell])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains('PRICED FOR PERFECTION')
      .setBackground(P.CHIP_NEG)
      .setFontColor(P.NEG_TXT)
      .setRanges([fundamentalCell])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains('ZOMBIE')
      .setBackground(P.CHIP_NEG)
      .setFontColor(P.NEG_TXT)
      .setRanges([fundamentalCell])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains('FAIR')
      .setBackground(P.CHIP_NEU)
      .setFontColor(P.TEXT)
      .setRanges([fundamentalCell])
      .build()
  ];
  
  // DECISION cell (B6) - updated position
  const decisionCell = REPORT.getRange('B6');
  const decisionRules = [
    // Positive signals (Green)
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains('BREAKOUT BUY')
      .setBackground(P.CHIP_POS)
      .setFontColor(P.POS_TXT)
      .setRanges([decisionCell])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains('OVERSOLD BUY')
      .setBackground(P.CHIP_POS)
      .setFontColor(P.POS_TXT)
      .setRanges([decisionCell])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains('Strong Trade Long')
      .setBackground(P.CHIP_POS)
      .setFontColor(P.POS_TXT)
      .setRanges([decisionCell])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains('STRONG BUY')
      .setBackground(P.CHIP_POS)
      .setFontColor(P.POS_TXT)
      .setRanges([decisionCell])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains('BUY')
      .setBackground(P.CHIP_POS)
      .setFontColor(P.POS_TXT)
      .setRanges([decisionCell])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains('ACCUMULATE')
      .setBackground(P.CHIP_POS)
      .setFontColor(P.POS_TXT)
      .setRanges([decisionCell])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains('Trade Long')
      .setBackground(P.CHIP_POS)
      .setFontColor(P.POS_TXT)
      .setRanges([decisionCell])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains('Add in Dip')
      .setBackground(P.CHIP_POS)
      .setFontColor(P.POS_TXT)
      .setRanges([decisionCell])
      .build(),
    
    // Neutral signals (Gray)
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains('Wait for Breakout')
      .setBackground(P.CHIP_NEU)
      .setFontColor(P.TEXT)
      .setRanges([decisionCell])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains('HOLD')
      .setBackground(P.CHIP_NEU)
      .setFontColor(P.TEXT)
      .setRanges([decisionCell])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains('NEUTRAL')
      .setBackground(P.CHIP_NEU)
      .setFontColor(P.TEXT)
      .setRanges([decisionCell])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains('WATCH')
      .setBackground(P.CHIP_NEU)
      .setFontColor(P.TEXT)
      .setRanges([decisionCell])
      .build(),
    
    // Warning signals (Orange)
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains('OVERBOUGHT')
      .setBackground(P.CHIP_WARN)
      .setFontColor(P.WARN_TXT)
      .setRanges([decisionCell])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains('OVERSOLD')
      .setBackground(P.CHIP_WARN)
      .setFontColor(P.WARN_TXT)
      .setRanges([decisionCell])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains('Take Profit')
      .setBackground(P.CHIP_WARN)
      .setFontColor(P.WARN_TXT)
      .setRanges([decisionCell])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains('TRIM')
      .setBackground(P.CHIP_WARN)
      .setFontColor(P.WARN_TXT)
      .setRanges([decisionCell])
      .build(),
    
    // Negative signals (Red)
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains('STOP OUT')
      .setBackground(P.CHIP_NEG)
      .setFontColor(P.NEG_TXT)
      .setRanges([decisionCell])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains('RISK OFF')
      .setBackground(P.CHIP_NEG)
      .setFontColor(P.NEG_TXT)
      .setRanges([decisionCell])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains('Stop-Out')
      .setBackground(P.CHIP_NEG)
      .setFontColor(P.NEG_TXT)
      .setRanges([decisionCell])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains('AVOID')
      .setBackground(P.CHIP_NEG)
      .setFontColor(P.NEG_TXT)
      .setRanges([decisionCell])
      .build()
  ];
  
  // Apply all rules
  const sheet = REPORT;
  const existingRules = sheet.getConditionalFormatRules();
  const allNewRules = signalRules.concat(fundamentalRules).concat(decisionRules);
  sheet.setConditionalFormatRules(existingRules.concat(allNewRules));
}

/**
 * Apply conditional formatting to M.PRICE (B8)
 * Green if M.PRICE > Current Price, Red otherwise
 * Uses helper column to avoid cross-sheet reference error
 */
function applyMarketPriceConditionalFormatting_(REPORT) {
  const P = reportPalette___();
  
  // Get locale separator
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const locale = (ss.getSpreadsheetLocale() || "").toLowerCase();
  const SEP = (/^(en|en_)/.test(locale)) ? "," : ";";
  
  // M.PRICE cell (B8)
  const marketPriceCell = REPORT.getRange('B8');
  
  // Create helper formula in hidden column AX (column 50) to get current price
  const helperCol = 50; // Column AX
  const helperCell = REPORT.getRange(8, helperCol); // AX8
  helperCell.setFormula(`=IFERROR(VALUE(INDEX(CALCULATIONS!G:G${SEP}MATCH(UPPER(TRIM($A$1))${SEP}ARRAYFORMULA(UPPER(TRIM(CALCULATIONS!A:A)))${SEP}0)))${SEP}0)`);
  
  // Green if M.PRICE (B8) > Current Price (AX8)
  const greenRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(`=AND(NOT(ISBLANK(B8))${SEP}B8<>"—"${SEP}ISNUMBER(VALUE(B8))${SEP}VALUE(B8)>AX8${SEP}AX8>0)`)
    .setBackground(P.CHIP_POS)
    .setFontColor(P.POS_TXT)
    .setRanges([marketPriceCell])
    .build();
  
  // Red if M.PRICE (B8) <= Current Price (AX8)
  const redRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(`=AND(NOT(ISBLANK(B8))${SEP}B8<>"—"${SEP}ISNUMBER(VALUE(B8))${SEP}VALUE(B8)<=AX8${SEP}AX8>0)`)
    .setBackground(P.CHIP_NEG)
    .setFontColor(P.NEG_TXT)
    .setRanges([marketPriceCell])
    .build();
  
  // Apply rules
  const sheet = REPORT;
  const existingRules = sheet.getConditionalFormatRules();
  sheet.setConditionalFormatRules(existingRules.concat([greenRule, redRule]));
}

/**
 
* Apply SMA color coding using helper formulas
 */
function applySMAColorCoding_(REPORT, row, label) {
  const P = reportPalette___();
  const valueCell = REPORT.getRange(row, 2);
  
  // Create a helper formula in a hidden column to compare price vs SMA
  const helperCol = 50; // Column AX (far beyond visible range, not column D)
  let helperFormula = '';
  
  // CORRECT COLUMN REFERENCES per generateCalculations.js:
  // G=Price (col 7), O=SMA 20 (col 15), P=SMA 50 (col 16), Q=SMA 200 (col 17)
  // Logic: Green when Price >= SMA (bullish), Red when Price < SMA (bearish)
  
  switch (label) {
    case 'SMA 20':
      // Compare Price (G) >= SMA 20 (O)
      helperFormula = '=IF(IFERROR(INDEX(CALCULATIONS!G:G,MATCH(UPPER(TRIM($A$1)),ARRAYFORMULA(UPPER(TRIM(CALCULATIONS!A:A))),0)),0)>=IFERROR(INDEX(CALCULATIONS!O:O,MATCH(UPPER(TRIM($A$1)),ARRAYFORMULA(UPPER(TRIM(CALCULATIONS!A:A))),0)),0),1,0)';
      break;
    case 'SMA 50':
      // Compare Price (G) >= SMA 50 (P)
      helperFormula = '=IF(IFERROR(INDEX(CALCULATIONS!G:G,MATCH(UPPER(TRIM($A$1)),ARRAYFORMULA(UPPER(TRIM(CALCULATIONS!A:A))),0)),0)>=IFERROR(INDEX(CALCULATIONS!P:P,MATCH(UPPER(TRIM($A$1)),ARRAYFORMULA(UPPER(TRIM(CALCULATIONS!A:A))),0)),0),1,0)';
      break;
    case 'SMA 200':
      // Compare Price (G) >= SMA 200 (Q)
      helperFormula = '=IF(IFERROR(INDEX(CALCULATIONS!G:G,MATCH(UPPER(TRIM($A$1)),ARRAYFORMULA(UPPER(TRIM(CALCULATIONS!A:A))),0)),0)>=IFERROR(INDEX(CALCULATIONS!Q:Q,MATCH(UPPER(TRIM($A$1)),ARRAYFORMULA(UPPER(TRIM(CALCULATIONS!A:A))),0)),0),1,0)';
      break;
  }
  
  // Set helper formula in column AX (column 50)
  REPORT.getRange(row, helperCol).setFormula(helperFormula);
  
  // Hide column AX
  REPORT.hideColumns(helperCol);
  
  // Create conditional formatting rules based on helper column - use $AX for column 50
  const rules = [
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=$AX${row}=1`)
      .setBackground(P.CHIP_POS)
      .setFontColor(P.POS_TXT)
      .setRanges([valueCell])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=$AX${row}=0`)
      .setBackground(P.CHIP_NEG)
      .setFontColor(P.NEG_TXT)
      .setRanges([valueCell])
      .build()
  ];
  
  // Apply the rules
  const sheet = valueCell.getSheet();
  const existingRules = sheet.getConditionalFormatRules();
  sheet.setConditionalFormatRules(existingRules.concat(rules));
}/*
*
 * Apply Support/Resistance color coding using helper formulas
 */
function applySupportResistanceColorCoding_(REPORT, row, label) {
  const P = reportPalette___();
  const valueCell = REPORT.getRange(row, 2);
  
  // Create a helper formula in a hidden column to compare price vs Support/Resistance
  const helperCol = 50; // Column AX (far beyond visible range, not column D)
  let helperFormula = '';
  
  switch (label) {
    case 'Support':
      // Return 1 if price >= support (good), 0 if price < support (bad - stop out risk)
      helperFormula = '=IF(IFERROR(INDEX(CALCULATIONS!E:E,MATCH(UPPER(TRIM($A$1)),ARRAYFORMULA(UPPER(TRIM(CALCULATIONS!A:A))),0)),0)>=IFERROR(INDEX(CALCULATIONS!U:U,MATCH(UPPER(TRIM($A$1)),ARRAYFORMULA(UPPER(TRIM(CALCULATIONS!A:A))),0)),0),1,0)';
      break;
    case 'Resistance':
      // Return 1 if price > resistance (breakout - good), 0 if price <= resistance (neutral)
      helperFormula = '=IF(IFERROR(INDEX(CALCULATIONS!E:E,MATCH(UPPER(TRIM($A$1)),ARRAYFORMULA(UPPER(TRIM(CALCULATIONS!A:A))),0)),0)>IFERROR(INDEX(CALCULATIONS!V:V,MATCH(UPPER(TRIM($A$1)),ARRAYFORMULA(UPPER(TRIM(CALCULATIONS!A:A))),0)),0),1,0)';
      break;
  }
  
  // Set helper formula in column AX (column 50)
  REPORT.getRange(row, helperCol).setFormula(helperFormula);
  
  // Hide column AX
  REPORT.hideColumns(helperCol);
  
  // Create conditional formatting rules based on helper column - use $AX for column 50
  const rules = [];
  
  if (label === 'Support') {
    // Support: Red ONLY if below support, no color if above support
    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied(`=$AX${row}=0`)
        .setBackground(P.CHIP_NEG)
        .setFontColor(P.NEG_TXT)
        .setRanges([valueCell])
        .build()
      // No green for above support - just red for below support
    );
  } else if (label === 'Resistance') {
    // Resistance: Green ONLY if above resistance, no color if below resistance
    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied(`=$AX${row}=1`)
        .setBackground(P.CHIP_POS)
        .setFontColor(P.POS_TXT)
        .setRanges([valueCell])
        .build()
      // No color for below resistance - just green for above resistance
    );
  }
  
  // Apply the rules
  const sheet = valueCell.getSheet();
  const existingRules = sheet.getConditionalFormatRules();
  sheet.setConditionalFormatRules(existingRules.concat(rules));
}

/**
 * Setup REPORT sheet layout with cell merging and formatting
 * Requirements: 6.1, 6.2, 8.1, 8.2, 8.3, 9.1, 9.2
 */
function setupReportLayout(REPORT) {
  const P = reportPalette___();
  
  // Set professional dark background and white font for cells M1:N1 (was N1:O1 before column D deletion)
  REPORT.getRange('M1:N1')
    .setBackground('#212121')
    .setFontColor('#FFFFFF')
    .setFontWeight('normal')
    .setFontSize(9)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  
  // Merge cells A43:M65 with formula preservation
  // First, unmerge any existing merges in this range to avoid conflicts
  try {
    REPORT.getRange('A43:M65').breakApart();
  } catch (e) {
    // Range might not be merged, continue
  }
  
  const mergeRange1 = REPORT.getRange('A43:M65');
  const existingFormula1 = REPORT.getRange('A43').getFormula();
  mergeRange1.merge();
  if (existingFormula1) {
    REPORT.getRange('A43').setFormula(existingFormula1);
  }
}

/**
 * Position REPORT sheet chart at D3:N17 (AFTER column D deletion, was E3:O17 BEFORE)
 * Requirements: 7.1, 7.2
 */
function positionReportChart(chart) {
  return chart.modify()
    .setPosition(3, 4, 0, 0)  // Row 3, Column D
    .setOption('width', 880)   // Width to span D-N
    .setOption('height', 420)  // Height to span rows 3-17
    .build();
}

/**
 * Custom function for REPORT sheet institutional analysis
 * Usage: =DASH_REPORT(A1)
 * @param {string} ticker - The ticker symbol from cell A1
 * @return {string} Institutional analysis text
 * @customfunction
 */
function DASH_REPORT(ticker) {
  if (!ticker || ticker === "") return "";
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const CALC = ss.getSheetByName('CALCULATIONS');
  
  if (!CALC) return "CALCULATIONS sheet not found";
  
  // Get all data from CALCULATIONS sheet
  const calcData = CALC.getDataRange().getValues();
  
  // Find the ticker row
  const tickerUpper = String(ticker).toUpperCase().trim();
  const tickerRow = calcData.findIndex(row => String(row[0]).toUpperCase().trim() === tickerUpper);
  
  if (tickerRow === -1) return "Ticker not found in CALCULATIONS";
  
  const row = calcData[tickerRow];
  
  // Extract values (0-based indices)
  const decision = row[2] || "—";  // Column C
  const signal = row[3] || "—";    // Column D
  const patterns = row[4] || "—";  // Column E
  const price = Number(row[6]) || 0;  // Column G
  const rsi = Number(row[17]) || 0;   // Column R
  const adx = Number(row[20]) || 0;   // Column U
  const volTrend = Number(row[8]) || 0;  // Column I
  const sma50 = Number(row[15]) || 0;    // Column P
  const sma200 = Number(row[16]) || 0;   // Column Q
  const support = Number(row[28]) || 0;  // Column AC
  const resistance = Number(row[29]) || 0; // Column AD
  const bbp = Number(row[25]) || 0;      // Column Z
  const athDiff = Number(row[10]) || 0;  // Column K
  
  // Check if data is loaded
  if (price === 0) return "LOADING...";
  
  // Build the analysis text
  let text = "📊 INSTITUTIONAL ANALYSIS: " + ticker + "\n\n";
  text += "DECISION: " + decision + " | SIGNAL: " + signal + "\n";
  text += "PATTERNS: " + patterns + "\n";
  text += "═══════════════════════════════════════════════════\n\n";
  text += "🎯 WHY '" + signal + "' TRIGGERED:\n\n";
  
  // Signal-specific explanations
  if (signal === "ACCUMULATE") {
    text += "✅ Price ($" + price.toFixed(2) + ") > SMA 200 ($" + sma200.toFixed(2) + ") → Bullish long-term trend\n";
    text += "✅ RSI (" + rsi.toFixed(1) + ") between 35-55 → Neutral zone, not overbought/oversold\n";
    text += "✅ Price within 5% of SMA 50 → $" + price.toFixed(2) + " is within $" + sma50.toFixed(2) + " ± 5% range ($" + (sma50 * 0.95).toFixed(2) + " - $" + (sma50 * 1.05).toFixed(2) + ")\n\n";
    text += "This is a conservative accumulation zone - stock is in bullish trend but not overextended, making it safe to build position gradually.";
  } else if (signal === "STRONG BUY") {
    text += "✅ Price > SMA 200 → Bullish regime | SMA 50 > SMA 200 → Uptrend confirmed\n";
    text += "✅ RSI (" + rsi.toFixed(1) + ") 30-40 → Early entry | MACD > 0 → Positive momentum\n";
    text += "✅ ADX (" + adx.toFixed(1) + ") ≥ 20 → Strong trend | Vol (" + volTrend.toFixed(2) + "x) ≥ 1.5x → High participation\n\n";
    text += "High-conviction entry with strong trend, momentum, and volume confirmation.";
  } else if (signal === "BUY") {
    text += "✅ Price > SMA 200 → Bullish | SMA 50 > SMA 200 → Uptrend intact\n";
    text += "✅ RSI (" + rsi.toFixed(1) + ") 40-50 → Healthy pullback | MACD > 0 → Momentum positive\n";
    text += "✅ ADX (" + adx.toFixed(1) + ") ≥ 15 → Developing trend\n\n";
    text += "Standard buy signal with good risk/reward. Trend established but not overextended.";
  } else if (signal === "STOP OUT") {
    text += "🔴 Price ($" + price.toFixed(2) + ") < Support ($" + support.toFixed(2) + ") → Support breakdown\n";
    text += "🔴 Risk management triggered → Exit position immediately to preserve capital\n\n";
    text += "Critical support breached. Technical structure compromised - exit to limit losses.";
  } else if (signal === "RISK OFF") {
    text += "🔴 Price ($" + price.toFixed(2) + ") < SMA 200 ($" + sma200.toFixed(2) + ") → Bear market regime\n";
    text += "🔴 Long-term trend broken → Avoid new long positions\n\n";
    text += "Bearish market structure. Stay defensive and wait for trend reversal confirmation.";
  } else if (signal === "TRIM") {
    text += "⚠️ RSI (" + rsi.toFixed(1) + ") ≥ 70 → Overbought | Bollinger %B (" + (bbp * 100).toFixed(1) + "%) ≥ 85% → Extended\n";
    text += "⚠️ Price near Resistance ($" + resistance.toFixed(2) + ") → Supply zone\n\n";
    text += "Stock overextended. Consider taking partial profits to lock in gains.";
  } else {
    text += "Signal criteria: " + signal + " - see formula logic for details";
  }
  
  text += "\n\n─────────────────────────────────────────────────────\n📉 OTHER INDICATOR ANALYSIS:\n\n";
  
  // RSI Analysis
  text += "• RSI (" + rsi.toFixed(1) + "): ";
  if (rsi >= 70) text += "Overbought (≥70) - momentum exhaustion risk";
  else if (rsi >= 60) text += "Elevated (60-70) - strong but approaching overbought";
  else if (rsi >= 50) text += "Bullish momentum (50-60) - healthy uptrend";
  else if (rsi >= 40) text += "Neutral (40-50) - balanced, waiting for bias";
  else if (rsi >= 30) text += "Pullback zone (30-40) - dip buying opportunity";
  else text += "Oversold (<30) - potential bounce, needs confirmation";
  text += "\n";
  
  // ADX Analysis
  text += "• ADX (" + adx.toFixed(1) + "): ";
  if (adx >= 25) text += "Strong trend (≥25)";
  else if (adx >= 20) text += "Developing trend (20-25)";
  else if (adx >= 15) text += "Weak trend (15-20)";
  else text += "No trend (<15) - range-bound";
  text += "\n";
  
  // Volume Trend Analysis
  text += "• Vol Trend (" + volTrend.toFixed(2) + "x): ";
  if (volTrend >= 2.0) text += "Extreme (≥2.0x) - institutional activity";
  else if (volTrend >= 1.5) text += "Strong (1.5-2.0x) - above average";
  else if (volTrend >= 1.0) text += "Normal (1.0-1.5x) - average";
  else if (volTrend >= 0.7) text += "Low (0.7-1.0x) - drift risk";
  else text += "Very low (<0.7x) - lack of conviction";
  text += "\n";
  
  // Bollinger %B Analysis
  text += "• Bollinger %B (" + (bbp * 100).toFixed(1) + "%): ";
  if (bbp >= 0.85) text += "Overbought (≥85%) - extended, limited upside";
  else if (bbp >= 0.5) text += "Upper half (50-85%) - bullish bias";
  else if (bbp >= 0.2) text += "Lower half (20-50%) - neutral to bearish";
  else text += "Oversold (<20%) - near lower band, bounce potential";
  text += "\n";
  
  // Price Position Analysis
  text += "• Price Position: ";
  if (price < support) text += "Below Support ($" + support.toFixed(2) + ") - breakdown";
  else if (price < sma200) text += "Below SMA 200 - bear regime";
  else if (price > resistance * 0.98) text += "Near Resistance ($" + resistance.toFixed(2) + ") - supply zone";
  else text += "Mid-range - neutral positioning";
  text += "\n";
  
  // ATH Diff Analysis
  text += "• ATH Diff (" + (athDiff * 100).toFixed(1) + "%): ";
  if (athDiff >= -0.02) text += "At ATH (≥-2%) - market leader";
  else if (athDiff >= -0.05) text += "Near ATH (-2% to -5%)";
  else if (athDiff >= -0.15) text += "Pullback (-5% to -15%)";
  else if (athDiff >= -0.30) text += "Correction (-15% to -30%)";
  else text += "Deep value (<-30%)";
  text += "\n\n";
  
  text += "─────────────────────────────────────────────────────\n🎯 HOW DECISION WAS DERIVED:\n\n";
  text += "1. SIGNAL: " + signal + " (technical setup)\n";
  text += "2. PATTERNS: " + (patterns === "—" ? "not detected" : patterns + " detected") + "\n";
  if (patterns !== "—") {
    text += "   Bullish patterns (ASC_TRI, BRKOUT, DBL_BTM, INV_H&S, CUP_HDL) confirm longs\n";
    text += "   Bearish patterns (DESC_TRI, H&S, DBL_TOP) create conflicts\n";
  }
  text += "3. DECISION: " + decision + "\n";
  
  // Decision explanation
  if (decision.includes("PATTERN CONFIRMED")) {
    text += "   ✅ " + signal + " + Bullish pattern = HIGH-CONFIDENCE ENTRY";
  } else if (decision.includes("PATTERN CONFLICT")) {
    text += "   ⚠️ " + signal + " BUT bearish pattern = WAIT FOR CLARITY";
  } else if (decision.includes("BUY") || decision.includes("ADD")) {
    text += "   ✅ Positive signal supports entry/add";
  } else if (decision.includes("EXIT")) {
    text += "   🔴 STOP OUT/RISK OFF = IMMEDIATE EXIT";
  } else if (decision.includes("TRIM")) {
    text += "   🟠 Overbought = TAKE PARTIAL PROFITS";
  } else if (decision.includes("HOLD")) {
    text += "   ⚖️ No actionable signal = MAINTAIN POSITION";
  } else {
    text += "   ⚪ Neutral - wait for clearer setup";
  }
  
  text += "\n\n📌 NOTE: FUNDAMENTAL data is informational only. Technical analysis drives all signals.";
  
  return text;
}