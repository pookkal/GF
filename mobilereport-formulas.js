/**
 * ==============================================================================
 * CLEAN MOBILE REPORT GENERATOR - GOOGLE SHEETS FORMULA VERSION
 * ==============================================================================
 * Run: setupFormulaBasedReport()
 */

/**
 * MAIN ENTRY POINT - Run this function
 */
function setupFormulaBasedReport() {
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
  
  // Create initial chart if ticker is selected
  const ticker = String(REPORT.getRange('A1').getValue() || '').trim();
  if (ticker) {
    createReportChart_(REPORT);
  }
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
  
  // Set professional font for entire sheet
  const maxRows = Math.max(50, REPORT.getLastRow());
  REPORT.getRange(1, 1, maxRows, 12).setFontFamily('Calibri');
  
  // Helper function for robust lookups
  const lookup = (col) => `=IFERROR(INDEX(CALCULATIONS!${col}:${col},MATCH(UPPER(TRIM($A$1)),ARRAYFORMULA(UPPER(TRIM(CALCULATIONS!A:A))),0)),"—")`;
  
  // Numeric lookup for calculations
  const numLookup = (col) => `IFERROR(VALUE(INDEX(CALCULATIONS!${col}:${col},MATCH(UPPER(TRIM($A$1)),ARRAYFORMULA(UPPER(TRIM(CALCULATIONS!A:A))),0))),0)`;
  
  // Row 1: Ticker name merged A1:C1, Date in D1
  REPORT.getRange('A1:C1').merge()
    .setBackground(P.BG_TOP)
    .setFontColor('#FFFFFF')
    .setFontWeight('normal')
    .setFontSize(14)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  
  // Set ticker dropdown validation on the merged cell A1:C1
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const INPUT = ss.getSheetByName('INPUT');
  if (INPUT) {
    setupReportTickerDropdown_(REPORT, INPUT);
  }
  
  // Date in D1 (calculated from dropdowns)
  REPORT.getRange('D1').setFormula('=IF(AND(ISNUMBER(VALUE(LEFT(A2,LEN(A2)-1))),ISNUMBER(VALUE(LEFT(B2,LEN(B2)-1))),ISNUMBER(VALUE(LEFT(C2,LEN(C2)-1)))),TEXT(TODAY()-VALUE(LEFT(A2,LEN(A2)-1))*365-VALUE(LEFT(B2,LEN(B2)-1))*30-VALUE(LEFT(C2,LEN(C2)-1)),"yyyy-mm-dd"),"Select Date")')
    .setBackground(P.BG_TOP)
    .setFontColor('#FFFFFF')
    .setFontWeight('normal')
    .setFontSize(12)
    .setFontFamily('Calibri')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  
  // Row 2: Date selection dropdowns A2:C2, Interval dropdown D2
  // Years dropdown (A2): 0Y to 20Y
  const yearsValues = [];
  for (let i = 0; i <= 20; i++) {
    yearsValues.push([i + 'Y']);
  }
  REPORT.getRange('A2')
    .setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(yearsValues.flat(), true).build())
    .setValue('0Y')
    .setBackground(P.BG_TOP)
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
    .setBackground(P.BG_TOP)
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
    .setBackground(P.BG_TOP)
    .setFontColor('#FFFFFF')
    .setFontWeight('normal')
    .setFontSize(10)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  
  // Row 3: NEW ROW - Calculated date display (A3:B3 merged) and Weekly/Daily dropdown (C3)
  // Calculated date display (A3:B3 merged)
  REPORT.getRange('A3:B3').merge()
    .setFormula('=IF(AND(ISNUMBER(VALUE(LEFT(A2,LEN(A2)-1))),ISNUMBER(VALUE(LEFT(B2,LEN(B2)-1))),ISNUMBER(VALUE(LEFT(C2,LEN(C2)-1)))),TEXT(TODAY()-VALUE(LEFT(A2,LEN(A2)-1))*365-VALUE(LEFT(B2,LEN(B2)-1))*30-VALUE(LEFT(C2,LEN(C2)-1)),"yyyy-mm-dd"),"Select Date")')
    .setBackground(P.BG_TOP)
    .setFontColor('#FFFFFF')
    .setFontWeight('normal')
    .setFontSize(12)
    .setFontFamily('Calibri')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  
  // Weekly/Daily dropdown moved to C3
  REPORT.getRange('C3')
    .setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Weekly', 'Daily'], true).build())
    .setValue('Daily')
    .setBackground(P.BG_TOP)
    .setFontColor('#FFFFFF')
    .setFontWeight('normal')
    .setFontSize(10)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  
  // Chart controls starting from E1:N2 (labels in row 1, checkboxes in row 2)
  setupChartControls_(REPORT);
  
  // Decision section starts at row 4 (moved down to accommodate new row 3)
  REPORT.getRange('A4').setValue('SIGNAL');
  REPORT.getRange('B4').setFormula(lookup('B'));
  REPORT.getRange('B4:D4').merge();
  
  REPORT.getRange('A5').setValue('FUNDAMENTAL');
  REPORT.getRange('B5').setFormula(lookup('C')); // FUNDAMENTAL from CALCULATIONS column C
  REPORT.getRange('B5:D5').merge();
  
  REPORT.getRange('A6').setValue('DECISION');
  REPORT.getRange('B6').setFormula(lookup('D')); // DECISION from CALCULATIONS column D
  REPORT.getRange('B6:D6').merge();
  
  // Style decision section
  REPORT.getRange('A4:D6')
    .setFontColor(P.BLACK)
    .setFontWeight('normal')
    .setFontSize(11)
    .setFontFamily('Calibri')
    .setBorder(true, true, true, true, true, true, P.GRID, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  
  // Apply yellow background only to labels (column A)
  REPORT.getRange('A4:A6')
    .setBackground(P.YELLOW);
  
  // Set decision value cells to dark background
  REPORT.getRange('B4:D6')
    .setBackground(P.BG_ROW_A)
    .setFontColor(P.TEXT);
  
  // Apply conditional formatting to decision cells
  applyDecisionConditionalFormatting_(REPORT);
  
  // Regime status (now row 7, was row 6)
  REPORT.getRange('A7').setFormula(`=IF(ISBLANK($A$1),"Select ticker in A1",IF(IFERROR(INDEX(CALCULATIONS!E:E,MATCH(UPPER(TRIM($A$1)),ARRAYFORMULA(UPPER(TRIM(CALCULATIONS!A:A))),0)),0)>=IFERROR(INDEX(CALCULATIONS!O:O,MATCH(UPPER(TRIM($A$1)),ARRAYFORMULA(UPPER(TRIM(CALCULATIONS!A:A))),0)),0),"RISK-ON (Above SMA200)","RISK-OFF (Below SMA200)"))`);
  REPORT.getRange('A7:D7').merge()
    .setBackground('#1F2937')
    .setFontColor('#FFFFFF')
    .setFontWeight('normal')
    .setFontFamily('Calibri');
  
  // Chart section (E3:M22) - starts after controls
  setupChartSection_(REPORT);
  
  // Data rows start at row 8 (removed empty row, moved up from row 9)
  let row = 8;
  
  // MARKET SNAPSHOT Section - Basic Price & Fundamental Data (Added PRICE)
  row = addSection_(REPORT, row, 'MARKET SNAPSHOT');
  row = addDataRow_(REPORT, row, 'PRICE', lookup('E'), '$#,##0.00');
  row = addDataRow_(REPORT, row, 'CHG%', lookup('F'), '0.00%');
  row = addDataRow_(REPORT, row, 'P/E', '=IFERROR(GOOGLEFINANCE($A$1,"pe"),"")', '0.00');
  row = addDataRow_(REPORT, row, 'EPS', '=IFERROR(GOOGLEFINANCE($A$1,"eps"),"")', '0.00');
  row = addDataRow_(REPORT, row, 'ATH', lookup('H'), '$#,##0.00');
  row = addDataRow_(REPORT, row, 'ATH %', lookup('I'), '0.00%');
  row = addDataRow_(REPORT, row, 'Range %', `=IFERROR(IF(AND(ISNUMBER(VALUE(LEFT(A2,LEN(A2)-1))),ISNUMBER(VALUE(LEFT(B2,LEN(B2)-1))),ISNUMBER(VALUE(LEFT(C2,LEN(C2)-1)))),LET(currentPrice,GOOGLEFINANCE($A$1,"price"),historicalDate,TODAY()-VALUE(LEFT(A2,LEN(A2)-1))*365-VALUE(LEFT(B2,LEN(B2)-1))*30-VALUE(LEFT(C2,LEN(C2)-1)),historicalPrice,INDEX(GOOGLEFINANCE($A$1,"close",historicalDate,historicalDate+1),2,2),(currentPrice/historicalPrice-1)),"Select Date"),"—")`, '0.00%');
  
  // TREND ANALYSIS Section - Moving Averages & Trend Strength
  row = addSection_(REPORT, row, 'TREND ANALYSIS');
  row = addDataRow_(REPORT, row, 'SMA20', lookup('M'), '$#,##0.00');
  row = addDataRow_(REPORT, row, 'SMA50', lookup('N'), '$#,##0.00');
  row = addDataRow_(REPORT, row, 'SMA200', lookup('O'), '$#,##0.00');
  row = addDataRow_(REPORT, row, 'ADX', lookup('S'), '0.00');
  row = addDataRow_(REPORT, row, 'TREND STATE', lookup('L'), '@');
  row = addDataRow_(REPORT, row, 'TREND SCORE', lookup('K'), '@');
  
  // MOMENTUM OSCILLATORS Section - MACD, Stochastic (RSI removed as requested)
  row = addSection_(REPORT, row, 'MOMENTUM OSCILLATORS');
  row = addDataRow_(REPORT, row, 'MACD Hist', lookup('Q'), '0.000');
  row = addDataRow_(REPORT, row, 'Stoch %K', lookup('T'), '0.0%');
  row = addDataRow_(REPORT, row, 'Divergence', lookup('R'), '@');
  
  // VOLATILITY & VOLUME Section - ATR, Volume, Bollinger Bands
  row = addSection_(REPORT, row, 'VOLATILITY & VOLUME');
  row = addDataRow_(REPORT, row, 'ATR', lookup('X'), '0.00');
  row = addDataRow_(REPORT, row, 'RVOL', lookup('G'), '0.00"x"');
  row = addDataRow_(REPORT, row, 'Bollinger %B', lookup('Y'), '0.0%');
  row = addDataRow_(REPORT, row, 'VOL REGIME', lookup('AC'), '@');
  
  // SUPPORT & RESISTANCE Section - Key Levels & Targets
  row = addSection_(REPORT, row, 'SUPPORT & RESISTANCE');
  row = addDataRow_(REPORT, row, 'Support', lookup('U'), '$#,##0.00');
  row = addDataRow_(REPORT, row, 'Resistance', lookup('V'), '$#,##0.00');
  row = addDataRow_(REPORT, row, 'Target', lookup('W'), '$#,##0.00');
  row = addDataRow_(REPORT, row, 'R:R Ratio', lookup('J'), '0.00"x"');
  
  // ENHANCED PATTERNS Section - New Pattern Recognition
  row = addSection_(REPORT, row, 'ENHANCED PATTERNS');
  row = addDataRow_(REPORT, row, 'ATH ZONE', lookup('AD'), '@');
  row = addDataRow_(REPORT, row, 'BBP SIGNAL', lookup('AE'), '@');
  row = addDataRow_(REPORT, row, 'PATTERNS', lookup('AF'), '@');
  
  // RISK MANAGEMENT Section - ATR-based Stops & Targets
  row = addSection_(REPORT, row, 'RISK MANAGEMENT');
  row = addDataRow_(REPORT, row, 'ATR STOP', lookup('AG'), '$#,##0.00');
  row = addDataRow_(REPORT, row, 'ATR TARGET', lookup('AH'), '$#,##0.00');
  row = addDataRow_(REPORT, row, 'POSITION SIZE', lookup('Z'), '@');
  
  // Narrative sections - only FUND NOTES (now in column AB)
  row = addNarrative_(REPORT, row, 'FUND NOTES', `=IFERROR(INDEX(CALCULATIONS!AB:AB,MATCH(UPPER(TRIM($A$1)),ARRAYFORMULA(UPPER(TRIM(CALCULATIONS!A:A))),0)),"—")`);
  
  // Final styling
  const finalRow = REPORT.getLastRow();
  REPORT.getRange(1, 1, finalRow, 13) // Extended to column M for chart controls
    .setBorder(true, true, true, true, true, true, P.GRID, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  REPORT.setHiddenGridlines(true);
  
  // Create the chart
  createReportChart_(REPORT);
}

/**
 * Add section header
 */
function addSection_(REPORT, row, title) {
  const P = reportPalette___();
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
  return row + 1;
}

/**
 * Add data row with enhanced formatting for rows 35-45 (two columns with inferences)
 */
function addDataRow_(REPORT, row, label, formula, format) {
  const P = reportPalette___();
  
  // Check if this is in the enhanced inference section (rows 35-45)
  const isInferenceSection = (row >= 35 && row <= 45);
  
  // Label
  REPORT.getRange(row, 1).setValue(label);
  
  // Formula in column B
  REPORT.getRange(row, 2).setFormula(formula);
  
  // Enhanced two-column format for inference section (rows 35-45)
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
    // Original format for other rows - check if split zone (rows 8-34)
    const isSplit = (row >= 8 && row <= 34);
    if (isSplit) {
      const narrativeFormula = getNarrativeFormula_(label);
      REPORT.getRange(row, 3).setFormula(narrativeFormula);
    } else {
      // Merge B:C for non-split rows
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
  if (label === 'SMA20' || label === 'SMA50' || label === 'SMA200') {
    applySMAColorCoding_(REPORT, row, label);
  }
  
  // Special handling for Support/Resistance color coding
  if (label === 'Support' || label === 'Resistance') {
    applySupportResistanceColorCoding_(REPORT, row, label);
  }
  
  // Borders
  REPORT.getRange(row, 1, 1, 3)
    .setBorder(false, false, true, false, false, false, P.GRID, SpreadsheetApp.BorderStyle.SOLID);
  
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
    case 'PRICE':
      return '=IFERROR("Last price " & TEXT(' + lookup('E') + ',"$#,##0.00") & ".","—")';
    
    case 'CHG%':
      return '=IFERROR(IF(' + lookup('F') + '>0,"Up " & TEXT(ABS(' + lookup('F') + '),"0.00%") & " today.",IF(' + lookup('F') + '<0,"Down " & TEXT(ABS(' + lookup('F') + '),"0.00%") & " today.","Flat today.")),"—")';
    
    case 'P/E':
      return '=IFERROR("P/E ratio: " & TEXT(GOOGLEFINANCE($A$1,"pe"),"0.00") & IF(GOOGLEFINANCE($A$1,"pe")<=25," (attractive)",IF(GOOGLEFINANCE($A$1,"pe")<=35," (fair)",IF(GOOGLEFINANCE($A$1,"pe")<=60," (expensive)"," (extreme)"))),"P/E data unavailable")';
    
    case 'EPS':
      return '=IFERROR("Earnings per share: $" & TEXT(GOOGLEFINANCE($A$1,"eps"),"0.00") & IF(GOOGLEFINANCE($A$1,"eps")>=0.50," (profitable)",IF(GOOGLEFINANCE($A$1,"eps")>0," (weak)"," (losing money)")),"EPS data unavailable")';
    
    case 'ATH':
      return '=IFERROR("All-time high: " & TEXT(' + numLookup('H') + ',"$#,##0.00") & " (current price " & TEXT((' + numLookup('E') + '/' + numLookup('H') + '-1),"+0.0%;-0.0%") & " vs ATH).","—")';
    
    case 'ATH %':
      return '=IFERROR(TEXT(' + lookup('I') + ',"+0.00%;-0.00%") & " from ATH" & IF(' + lookup('I') + '>=-0.02," - at resistance zone.",IF(' + lookup('I') + '>=-0.15," - pullback zone."," - correction territory.")),"—")';
    
    case 'Range %':
      return '=IFERROR(IF(AND(ISNUMBER(VALUE(LEFT(A2,LEN(A2)-1))),ISNUMBER(VALUE(LEFT(B2,LEN(B2)-1))),ISNUMBER(VALUE(LEFT(C2,LEN(C2)-1)))),LET(currentPrice,GOOGLEFINANCE($A$1,"price"),historicalDate,TODAY()-VALUE(LEFT(A2,LEN(A2)-1))*365-VALUE(LEFT(B2,LEN(B2)-1))*30-VALUE(LEFT(C2,LEN(C2)-1)),historicalPrice,INDEX(GOOGLEFINANCE($A$1,"close",historicalDate,historicalDate+1),2,2),rangePercent,(currentPrice/historicalPrice-1),TEXT(rangePercent,"+0.00%;-0.00%") & " from " & TEXT(historicalDate,"yyyy-mm-dd") & " price " & TEXT(historicalPrice,"$0.00")),"Select date range first."),"—")';
    
    case 'SMA20':
      return '=IFERROR(TEXT((' + numLookup('E') + '/' + numLookup('M') + '-1),"+0.0%;-0.0%") & " vs SMA20 " & TEXT(' + numLookup('M') + ',"$0.00") & IF(' + numLookup('E') + '>=' + numLookup('M') + '," - short-term bullish."," - short-term bearish."),\"—\")';
    
    case 'SMA50':
      return '=IFERROR(TEXT((' + numLookup('E') + '/' + numLookup('N') + '-1),"+0.0%;-0.0%") & " vs SMA50 " & TEXT(' + numLookup('N') + ',"$0.00") & IF(' + numLookup('E') + '>=' + numLookup('N') + '," - medium-term bullish."," - medium-term bearish."),\"—\")';
    
    case 'SMA200':
      return '=IFERROR(TEXT((' + numLookup('E') + '/' + numLookup('O') + '-1),"+0.0%;-0.0%") & " vs SMA200 " & TEXT(' + numLookup('O') + ',"$0.00") & IF(' + numLookup('E') + '>=' + numLookup('O') + '," - RISK-ON regime."," - RISK-OFF regime."),\"—\")';
    
    case 'ADX':
      return '=IFERROR("Trend strength: " & TEXT(' + lookup('S') + ',"0.0") & IF(' + lookup('S') + '>=25," - strong trend.",IF(' + lookup('S') + '>=20," - trend developing.",IF(' + lookup('S') + '>=15," - weak trend."," - range-bound."))),"—")';
    
    case 'TREND STATE':
      return '=IFERROR(' + lookup('L') + ' & " market regime based on SMA200 position.","—")';
    
    case 'TREND SCORE':
      return '=IFERROR("Moving average alignment: " & ' + lookup('K') + ' & " stars (max 3 for price above all SMAs).","—")';
    
    case 'RSI':
      return '=IFERROR("RSI(14): " & TEXT(' + lookup('P') + ',"0.0") & IF(' + lookup('P') + '>=70," - overbought zone.",IF(' + lookup('P') + '<=30," - oversold zone.",IF(' + lookup('P') + '>=55," - positive momentum.",IF(' + lookup('P') + '<=45," - weak momentum."," - neutral range.")))),"—")';
    
    case 'MACD Hist':
      return '=IFERROR("MACD histogram: " & TEXT(' + lookup('Q') + ',"0.000") & IF(' + lookup('Q') + '>0," - positive momentum impulse.",IF(' + lookup('Q') + '<0," - negative momentum impulse."," - flat momentum.")),"—")';
    
    case 'Stoch %K':
      return '=IFERROR("Stochastic %K: " & TEXT(' + lookup('T') + ',"0.0%") & IF(' + lookup('T') + '>=0.8," - overbought timing.",IF(' + lookup('T') + '<=0.2," - oversold timing."," - neutral timing.")),"—")';
    
    case 'Divergence':
      return '=IFERROR(IF(' + lookup('R') + '="BULL DIV","Bullish divergence detected - price lower but momentum higher.",IF(' + lookup('R') + '="BEAR DIV","Bearish divergence detected - price higher but momentum lower.","No momentum divergence detected.")),"—")';
    
    case 'ATR':
      return '=IFERROR("Average True Range: " & TEXT(' + lookup('X') + ',"$0.00") & " (" & TEXT(' + lookup('X') + '/' + lookup('E') + ',"0.0%") & " of price) - volatility measure.","—")';
    
    case 'RVOL':
      return '=IFERROR("Relative volume: " & TEXT(' + lookup('G') + ',"0.00") & "x average" & IF(' + lookup('G') + '>=1.5," - strong participation.",IF(' + lookup('G') + '>=1," - average participation."," - low participation (drift risk).")),"—")';
    
    case 'Bollinger %B':
      return '=IFERROR("Bollinger %B: " & TEXT(' + lookup('Y') + ',"0.0%") & IF(' + lookup('Y') + '>1," - above upper band (expansion).",IF(' + lookup('Y') + '>=0.8," - upper band zone.",IF(' + lookup('Y') + '<0," - below lower band (extreme).",IF(' + lookup('Y') + '<=0.2," - lower band zone."," - mid-band zone.")))),"—")';
    
    case 'VOL REGIME':
      return '=IFERROR(' + lookup('AC') + ' & " - volatility classification based on ATR/Price ratio for position sizing.","—")';
    
    case 'Support':
      return '=IFERROR("Support level: " & TEXT(' + numLookup('U') + ',"$0.00") & " (" & TEXT((' + numLookup('E') + '/' + numLookup('U') + '-1),"+0.0%;-0.0%") & ")" & IF(' + numLookup('E') + '<' + numLookup('U') + '," - BREAKDOWN RISK."," - holding above support."),\"—\")';
    
    case 'Resistance':
      return '=IFERROR("Resistance level: " & TEXT(' + numLookup('V') + ',"$0.00") & " (" & TEXT((' + numLookup('V') + '/' + numLookup('E') + '-1),"+0.0%;-0.0%") & " away)" & IF(' + numLookup('E') + '>' + numLookup('V') + '," - BREAKOUT CONFIRMED."," - below resistance."),\"—\")';
    
    case 'Target':
      return '=IFERROR("Price target: " & TEXT(' + numLookup('W') + ',"$#,##0.00") & " (" & TEXT((' + numLookup('W') + '/' + numLookup('E') + '-1),"+0.00%;-0.00%") & " upside potential).","—")';
    
    case 'R:R Ratio':
      return '=IFERROR("Risk/Reward ratio: " & TEXT(' + lookup('J') + ',"0.00") & ":1" & IF(' + lookup('J') + '>=3," - elite asymmetry.",IF(' + lookup('J') + '>=1.5," - acceptable asymmetry."," - poor asymmetry.")),"—")';
    
    case 'ATH ZONE':
      return '=IFERROR(' + lookup('AD') + ' & " - psychological zone based on distance from all-time highs.","—")';
    
    case 'BBP SIGNAL':
      return '=IFERROR(' + lookup('AE') + ' & " - Bollinger Band position signal for mean reversion opportunities.","—")';
    
    case 'PATTERNS':
      return '=IFERROR(IF(ISBLANK(' + lookup('AF') + '),"No enhanced patterns detected.",' + lookup('AF') + ' & " - institutional-grade pattern recognition."),"—")';
    
    case 'ATR STOP':
      return '=IFERROR("ATR-based stop: " & TEXT(' + numLookup('AG') + ',"$#,##0.00") & " (" & TEXT((' + numLookup('E') + '/' + numLookup('AG') + '-1),"+0.0%;-0.0%") & " risk from current price).","—")';
    
    case 'ATR TARGET':
      return '=IFERROR("ATR-based target: " & TEXT(' + numLookup('AH') + ',"$#,##0.00") & " (" & TEXT((' + numLookup('AH') + '/' + numLookup('E') + '-1),"+0.0%;-0.0%") & " reward potential).","—")';
    
    case 'POSITION SIZE':
      return '=IFERROR(' + lookup('Z') + ' & " - volatility and ATH-adjusted institutional position sizing.","—")';
    
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
      
    case 'ATH %':
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
      
    case 'RVOL':
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
      
    case 'R:R RATIO':
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
      
    case 'ADX':
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
      
    case 'STOCH %K':
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
      
    case 'SMA20':
    case 'SMA50':  
    case 'SMA200':
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
      .setAllowInvalid(false)
      .build();
    
    const a1 = reportSheet.getRange('A1');
    a1.setDataValidation(rule);
    
    if (!a1.getValue()) {
      a1.setValue('AAPL');
    }
    
    SpreadsheetApp.flush();
    Utilities.sleep(100);
    createReportChart_(reportSheet);
    return;
  }
  
  // Use DASHBOARD sheet A4:A range for ticker dropdown
  const dashboardLast = dashboardSheet.getLastRow();
  const dashboardHeight = Math.max(1, dashboardLast - 3); // Start from row 4, so subtract 3
  const dashboardRng = dashboardSheet.getRange(4, 1, dashboardHeight, 1); // A4:A range

  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(dashboardRng, true)
    .setAllowInvalid(false)
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
  
  // Force chart update when ticker is set
  SpreadsheetApp.flush();
  Utilities.sleep(100);
  createReportChart_(reportSheet);
}

/**
 * Setup chart control checkboxes - all in row 1 starting from E1
 */
/**
 * Setup chart control checkboxes - UPDATED VERSION without date controls (now in A2:D2)
 */
function setupChartControls_(REPORT) {
  // Clear row 1 and 2 from E to N first
  REPORT.getRange('E1:N2').clearContent().clearFormat();
  
  // All 9 controls in consecutive columns E through M (removed RSI)
  const controls = [
    ['PRICE', true],
    ['SMA20', false], 
    ['SMA50', false],
    ['SMA200', false],
    ['VOLUME', true],  // Enable volume by default to show bull/bear bars
    ['SUPPORT', true],
    ['RESISTANCE', true],
    ['ATR STOP', false],
    ['ATR TARGET', false]
  ];
  
  for (let i = 0; i < controls.length; i++) {
    const [label, defaultValue] = controls[i];
    const col = 5 + i; // E=5, F=6, G=7, H=8, I=9, J=10, K=11, L=12, M=13, N=14
    
    // Set column width to ensure all are visible
    REPORT.setColumnWidth(col, 80);
    
    // Add label in row 1
    REPORT.getRange(1, col).setValue(label);
    REPORT.getRange(1, col)
      .setBackground('#374151')
      .setFontColor('#FFFFFF')
      .setFontWeight('normal')
      .setFontSize(8)
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
    
    // Add checkbox in row 2
    REPORT.getRange(2, col).insertCheckboxes();
    REPORT.getRange(2, col).setValue(defaultValue);
    REPORT.getRange(2, col)
      .setBackground('#1F2937')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
  }
}

/**
 * Setup chart section placeholder - NO MERGE, prepare for floating chart (E3:M22)
 */
function setupChartSection_(REPORT) {
  // Clear the chart area completely - no merging, let chart float (E3:M22)
  REPORT.getRange('E3:M22').clearContent().clearFormat();
  
  // Just set a simple background color for the chart area
  REPORT.getRange('E3:M22')
    .setBackground('#0F172A')
    .setBorder(true, true, true, true, false, false, '#374151', SpreadsheetApp.BorderStyle.SOLID);
}

/**
 * Create dynamic chart using REPORT sheet data - PROPER FLOATING CHART WITH ENHANCED ERROR HANDLING
 */
function createReportChart_(REPORT) {
  try {
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
    
    // Try to show a simple error message in the chart area
    try {
      REPORT.getRange('E3:M22').clearContent().clearFormat();
      REPORT.getRange('E3').setValue(`Chart Error: ${e.toString().substring(0, 100)}`);
      REPORT.getRange('E3:M22')
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
  
  // Get checkbox states from row 2, columns E-M (5-13) - RSI removed
  const checkboxes = {
    PRICE: REPORT.getRange(2, 5).getValue() || false,
    SMA20: REPORT.getRange(2, 6).getValue() || false,
    SMA50: REPORT.getRange(2, 7).getValue() || false,
    SMA200: REPORT.getRange(2, 8).getValue() || false,
    VOLUME: REPORT.getRange(2, 9).getValue() || false,
    SUPPORT: REPORT.getRange(2, 10).getValue() || false,
    RESISTANCE: REPORT.getRange(2, 11).getValue() || false,
    ATR_STOP: REPORT.getRange(2, 12).getValue() || false,
    ATR_TARGET: REPORT.getRange(2, 13).getValue() || false
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
  
  // Get current values from CALCULATIONS sheet for support/resistance/ATR
  const CALC = ss.getSheetByName('CALCULATIONS');
  let support = 0, resistance = 0, atr = 0, currentRSI = 50, currentPrice = 0;
  
  if (CALC) {
    const calcData = CALC.getDataRange().getValues();
    const tickerRow = calcData.findIndex(row => String(row[0]).toUpperCase().trim() === ticker.toUpperCase());
    if (tickerRow !== -1) {
      const calcRow = calcData[tickerRow];
      currentPrice = Number(calcRow[4]) || 0; // Column E - Current Price
      support = Number(calcRow[20]) || 0; // Column U - Support  
      resistance = Number(calcRow[21]) || 0; // Column V - Resistance
      atr = Number(calcRow[23]) || 0; // Column X - ATR
      currentRSI = Number(calcRow[15]) || 50; // Column P - RSI
    }
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
      
      // Calculate ATR-based levels using current ATR (static values - straight lines)
      const atrStop = atr > 0 ? currentPrice - (atr * 2) : currentPrice * 0.95;
      const atrTarget = atr > 0 ? currentPrice + (atr * 3) : currentPrice * 1.05;
      
      // Build data row based on selected checkboxes - FIXED ORDER TO PREVENT SHIFTING
      const dataRow = [d];
      
      // Always add price first if selected
      if (checkboxes.PRICE) dataRow.push(close);
      
      // Add SMAs next (before volume to prevent shifting)
      if (checkboxes.SMA20 && s20) dataRow.push(s20);
      if (checkboxes.SMA50 && s50) dataRow.push(s50);
      if (checkboxes.SMA200 && s200) dataRow.push(s200);
      
      // Add support/resistance levels
      if (checkboxes.RESISTANCE && resistance > 0) dataRow.push(resistance);
      if (checkboxes.SUPPORT && support > 0) dataRow.push(support);
      
      // Add ATR levels
      if (checkboxes.ATR_STOP) dataRow.push(atrStop);
      if (checkboxes.ATR_TARGET) dataRow.push(atrTarget);
      
      // Add volume LAST to prevent series shifting issues - SIMPLIFIED LOGIC
      if (checkboxes.VOLUME) {
        const prevClose = (i > 4) ? Number(raw[i - 1][4]) : close;
        // Split volume into bull/bear based on price movement
        // Bull volume (green) when close >= previous close
        // Bear volume (red) when close < previous close
        const isBullVolume = close >= prevClose;
        
        // Use original volume values - secondary axis will handle scaling
        const bullVol = isBullVolume ? vol : 0;
        const bearVol = isBullVolume ? 0 : vol;
        
        // Debug volume calculation for first few rows
        if (i < 7) {
          console.log(`Row ${i}: close=${close}, vol=${vol}, isBull=${isBullVolume}, bullVol=${bullVol}, bearVol=${bearVol}`);
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
      
      // Calculate ATR-based levels using current ATR (static values - straight lines)
      const atrStop = atr > 0 ? livePrice - (atr * 2) : livePrice * 0.95;
      const atrTarget = atr > 0 ? livePrice + (atr * 3) : livePrice * 1.05;
      
      // Build live data row based on selected checkboxes - SAME ORDER AS HISTORICAL DATA
      const liveDataRow = [today];
      
      // Always add price first if selected
      if (checkboxes.PRICE) liveDataRow.push(livePrice);
      
      // Add SMAs next (before volume to prevent shifting)
      if (checkboxes.SMA20 && liveS20) liveDataRow.push(liveS20);
      if (checkboxes.SMA50 && liveS50) liveDataRow.push(liveS50);
      if (checkboxes.SMA200 && liveS200) liveDataRow.push(liveS200);
      
      // Add support/resistance levels
      if (checkboxes.RESISTANCE && resistance > 0) liveDataRow.push(resistance);
      if (checkboxes.SUPPORT && support > 0) liveDataRow.push(support);
      
      // Add ATR levels
      if (checkboxes.ATR_STOP) liveDataRow.push(atrStop);
      if (checkboxes.ATR_TARGET) liveDataRow.push(atrTarget);
      
      // Add volume LAST - use proxy volume for today (same as updateDynamicChart)
      if (checkboxes.VOLUME) {
        // Get max volume from historical data for proxy calculation
        const allHistoricalVols = raw.slice(4).map(r => Number(r[5])).filter(v => isFinite(v) && v > 0);
        const maxHistVol = allHistoricalVols.length > 0 ? Math.max(...allHistoricalVols) : 1000000;
        const proxyVolume = maxHistVol * 0.5; // Use 50% of max historical volume as proxy
        
        const isBullVolume = livePrice >= lastHistClose;
        
        // Use proxy volume for today's data
        const bullVol = isBullVolume ? proxyVolume : 0;
        const bearVol = isBullVolume ? 0 : proxyVolume;
        
        console.log(`Live volume proxy: ${proxyVolume}, isBull=${isBullVolume}, bullVol=${bullVol}, bearVol=${bearVol}`);
        
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
      color: "#B71C1C", 
      lineDashStyle: [4, 4], 
      labelInLegend: "Resistance",
      targetAxisIndex: 0 // Explicitly assign to primary axis
    };
    seriesIndex++;
  }
  
  if (checkboxes.SUPPORT) {
    seriesConfig[seriesIndex] = { 
      type: "line", 
      color: "#00FF41", 
      lineDashStyle: [4, 4], 
      labelInLegend: "Support",
      targetAxisIndex: 0 // Explicitly assign to primary axis
    };
    seriesIndex++;
  }
  
  if (checkboxes.ATR_STOP) {
    seriesConfig[seriesIndex] = { 
      type: "line", 
      color: "#FF5722", 
      lineWidth: 2, 
      lineDashStyle: [2, 2], 
      labelInLegend: "ATR Stop",
      targetAxisIndex: 0 // Explicitly assign to primary axis
    };
    seriesIndex++;
  }
  
  if (checkboxes.ATR_TARGET) {
    seriesConfig[seriesIndex] = { 
      type: "line", 
      color: "#8BC34A", 
      lineWidth: 2, 
      lineDashStyle: [2, 2], 
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
      color: "#2E7D32", // Darker green like updateDynamicChart()
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
  
  // If no checkboxes are selected, show at least price
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
  try {
    const chart = REPORT.newChart()
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
      .setPosition(3, 5, 0, 0) // Row 3, Column E - FLOATING CHART
      .setOption("width", 720)  // Exact width to fit E3:M22 perfectly (9 columns * 80px)
      .setOption("height", 420) // Height to fit the area properly
      .build();
    
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
        .setPosition(3, 5, 0, 0)
        .setOption("width", 720)
        .setOption("height", 420)
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
 * Update REPORT sheet chart - called from Code.js onEdit trigger
 */
function updateReportChart() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const REPORT = ss.getSheetByName('REPORT');
  if (!REPORT) return;
  
  // Add small delay to ensure cell values are updated
  SpreadsheetApp.flush();
  Utilities.sleep(100);
  
  // Simply call the existing chart creation function
  createReportChart_(REPORT);
}

/**
 * Set column widths - updated for new layout
 */
function setReportColumnWidthsAndWrap___(REPORT) {
  const pxPerChar = 8;
  
  // Split column C into two parts while keeping A-D overall size same as original A-C
  const originalTotalWidth = Math.max(70, Math.round(8 * pxPerChar + 15)) + Math.max(70, Math.round(8 * pxPerChar + 15)) + Math.max(170, Math.round(27 * pxPerChar + 20));
  
  // Redistribute: A (ticker/dropdown), B (values), C (split from original C), D (date)
  const colA = Math.max(70, Math.round(8 * pxPerChar + 15)); // Keep A same
  const colB = Math.max(70, Math.round(8 * pxPerChar + 15)); // Keep B same  
  const colC = Math.max(120, Math.round(20 * pxPerChar + 15)); // Reduced from original C
  const colD = Math.max(80, Math.round(10 * pxPerChar + 15)); // New D column
  
  // Set main columns
  REPORT.setColumnWidth(1, colA);
  REPORT.setColumnWidth(2, colB);
  REPORT.setColumnWidth(3, colC);
  REPORT.setColumnWidth(4, colD);
  
  // Chart control columns (E-N) - handled in setupChartControls_
  
  const lastRow = Math.max(1, REPORT.getLastRow());
  REPORT.getRange(1, 1, lastRow, 13).setWrap(true); // Extended to column M
}

/**
 * Color palette
 */
function reportPalette___() {
  return {
    BG_TOP: '#0B0F14',
    PANEL:  '#111827',
    BG_ROW_A: '#0F172A',
    BG_ROW_B: '#111827',
    GRID: '#374151',
    TEXT: '#E5E7EB',
    MUTED: '#9CA3AF',
    POS_TXT:  '#34D399',
    NEG_TXT:  '#F87171',
    WARN_TXT: '#FBBF24',
    CHIP_POS:  '#06281F',
    CHIP_NEG:  '#2A0B0B',
    CHIP_WARN: '#2A1E05',
    CHIP_NEU:  '#0B1220',
    YELLOW: '#FDE047',
    BLACK:  '#111827'
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
}/**
 
* Apply SMA color coding using helper formulas
 */
function applySMAColorCoding_(REPORT, row, label) {
  const P = reportPalette___();
  const valueCell = REPORT.getRange(row, 2);
  
  // Create a helper formula in a hidden column to compare price vs SMA
  const helperCol = 4; // Column D (hidden)
  let helperFormula = '';
  
  switch (label) {
    case 'SMA20':
      helperFormula = '=IF(IFERROR(INDEX(CALCULATIONS!E:E,MATCH(UPPER(TRIM($A$1)),ARRAYFORMULA(UPPER(TRIM(CALCULATIONS!A:A))),0)),0)>=IFERROR(INDEX(CALCULATIONS!M:M,MATCH(UPPER(TRIM($A$1)),ARRAYFORMULA(UPPER(TRIM(CALCULATIONS!A:A))),0)),0),1,0)';
      break;
    case 'SMA50':
      helperFormula = '=IF(IFERROR(INDEX(CALCULATIONS!E:E,MATCH(UPPER(TRIM($A$1)),ARRAYFORMULA(UPPER(TRIM(CALCULATIONS!A:A))),0)),0)>=IFERROR(INDEX(CALCULATIONS!N:N,MATCH(UPPER(TRIM($A$1)),ARRAYFORMULA(UPPER(TRIM(CALCULATIONS!A:A))),0)),0),1,0)';
      break;
    case 'SMA200':
      helperFormula = '=IF(IFERROR(INDEX(CALCULATIONS!E:E,MATCH(UPPER(TRIM($A$1)),ARRAYFORMULA(UPPER(TRIM(CALCULATIONS!A:A))),0)),0)>=IFERROR(INDEX(CALCULATIONS!O:O,MATCH(UPPER(TRIM($A$1)),ARRAYFORMULA(UPPER(TRIM(CALCULATIONS!A:A))),0)),0),1,0)';
      break;
  }
  
  // Set helper formula in column D
  REPORT.getRange(row, helperCol).setFormula(helperFormula);
  
  // Hide column D
  REPORT.hideColumns(helperCol);
  
  // Create conditional formatting rules based on helper column
  const rules = [
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=$D${row}=1`)
      .setBackground(P.CHIP_POS)
      .setFontColor(P.POS_TXT)
      .setRanges([valueCell])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=$D${row}=0`)
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
  const helperCol = 4; // Column D (hidden)
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
  
  // Set helper formula in column D
  REPORT.getRange(row, helperCol).setFormula(helperFormula);
  
  // Hide column D
  REPORT.hideColumns(helperCol);
  
  // Create conditional formatting rules based on helper column
  const rules = [];
  
  if (label === 'Support') {
    // Support: Red ONLY if below support, no color if above support
    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied(`=$D${row}=0`)
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
        .whenFormulaSatisfied(`=$D${row}=1`)
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