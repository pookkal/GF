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
  
  // Clear sheet
  REPORT.clear({ contentsOnly: true });
  REPORT.clearFormats();
  
  // Apply column widths
  setReportColumnWidthsAndWrap___(REPORT);
  
  // Create report
  createFormulaReport_(REPORT);
  
  SpreadsheetApp.flush();
}

/**
 * Create the complete formula-based report
 */
function createFormulaReport_(REPORT) {
  const P = reportPalette___();
  
  // Set professional font for entire sheet
  const maxRows = Math.max(50, REPORT.getLastRow());
  REPORT.getRange(1, 1, maxRows, 3).setFontFamily('Calibri');
  
  // Helper function for robust lookups
  const lookup = (col) => `=IFERROR(INDEX(CALCULATIONS!${col}:${col},MATCH(UPPER(TRIM($A$1)),ARRAYFORMULA(UPPER(TRIM(CALCULATIONS!A:A))),0)),"—")`;
  
  // Numeric lookup for calculations
  const numLookup = (col) => `IFERROR(VALUE(INDEX(CALCULATIONS!${col}:${col},MATCH(UPPER(TRIM($A$1)),ARRAYFORMULA(UPPER(TRIM(CALCULATIONS!A:A))),0))),0)`;
  
  // Header
  REPORT.getRange('A1').setValue(''); // Ticker dropdown
  REPORT.getRange('B1').setValue('MASTER REPORT');
  REPORT.getRange('B1:C1').merge();
  
  // Header styling
  REPORT.getRange('A1:C1')
    .setBackground(P.BG_TOP)
    .setFontColor('#FFFFFF')
    .setFontWeight('bold')
    .setFontSize(12)
    .setFontFamily('Calibri');
  
  // Decision section
  REPORT.getRange('A2').setValue('SIGNAL');
  REPORT.getRange('B2').setFormula(lookup('B'));
  REPORT.getRange('B2:C2').merge();
  
  REPORT.getRange('A3').setValue('FUNDAMENTAL');
  REPORT.getRange('B3').setFormula(lookup('D'));
  REPORT.getRange('B3:C3').merge();
  
  REPORT.getRange('A4').setValue('DECISION');
  REPORT.getRange('B4').setFormula(lookup('C'));
  REPORT.getRange('B4:C4').merge();
  
  // Style decision section (remove yellow background, keep black text)
  REPORT.getRange('A2:C4')
    .setFontColor(P.BLACK)
    .setFontWeight('bold')
    .setFontSize(11)
    .setFontFamily('Calibri')
    .setBorder(true, true, true, true, true, true, P.GRID, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  
  // Apply yellow background only to labels (column A)
  REPORT.getRange('A2:A4')
    .setBackground(P.YELLOW);
  
  // Set decision value cells to dark background
  REPORT.getRange('B2:C4')
    .setBackground(P.BG_ROW_A)
    .setFontColor(P.TEXT);
  
  // Apply conditional formatting to decision cells
  applyDecisionConditionalFormatting_(REPORT);
  
  // Timestamp
  REPORT.getRange('A5').setFormula('="Generated: " & TEXT(NOW(),"yyyy-mm-dd hh:mm")');
  REPORT.getRange('A5:C5').merge()
    .setBackground(P.BG_TOP)
    .setFontColor(P.MUTED)
    .setFontSize(9)
    .setFontFamily('Calibri');
  
  // Regime status (fixed formula)
  REPORT.getRange('A6').setFormula(`=IF(ISBLANK($A$1),"Select ticker in A1",IF(IFERROR(INDEX(CALCULATIONS!E:E,MATCH(UPPER(TRIM($A$1)),ARRAYFORMULA(UPPER(TRIM(CALCULATIONS!A:A))),0)),0)>=IFERROR(INDEX(CALCULATIONS!O:O,MATCH(UPPER(TRIM($A$1)),ARRAYFORMULA(UPPER(TRIM(CALCULATIONS!A:A))),0)),0),"RISK-ON (Above SMA200)","RISK-OFF (Below SMA200)"))`);
  REPORT.getRange('A6:C6').merge()
    .setBackground('#1F2937')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold')
    .setFontFamily('Calibri');
  
  // Data rows
  let row = 8;
  
  // SNAPSHOT Section
  row = addSection_(REPORT, row, 'SNAPSHOT');
  row = addDataRow_(REPORT, row, 'PRICE', lookup('E'), '$#,##0.00');
  row = addDataRow_(REPORT, row, 'CHG%', lookup('F'), '0.00%');
  row = addDataRow_(REPORT, row, 'RVOL', lookup('G'), '0.00"x"');
  row = addDataRow_(REPORT, row, 'TREND', `=IFERROR(INDEX(CALCULATIONS!L:L,MATCH(UPPER(TRIM($A$1)),ARRAYFORMULA(UPPER(TRIM(CALCULATIONS!A:A))),0)) & IF(ISBLANK(INDEX(CALCULATIONS!K:K,MATCH(UPPER(TRIM($A$1)),ARRAYFORMULA(UPPER(TRIM(CALCULATIONS!A:A))),0))),"",", Score: " & INDEX(CALCULATIONS!K:K,MATCH(UPPER(TRIM($A$1)),ARRAYFORMULA(UPPER(TRIM(CALCULATIONS!A:A))),0))),"—")`, '@');
  row = addDataRow_(REPORT, row, 'R:R', lookup('J'), '0.00"x"');
  row = addDataRow_(REPORT, row, 'P/E', '=IFERROR(GOOGLEFINANCE($A$1,"pe"),"")', '0.00');
  row = addDataRow_(REPORT, row, 'EPS', '=IFERROR(GOOGLEFINANCE($A$1,"eps"),"")', '0.00');
  
  // TREND & STRUCTURE Section
  row = addSection_(REPORT, row, 'TREND & STRUCTURE');
  row = addDataRow_(REPORT, row, 'SMA20', lookup('M'), '$#,##0.00');
  row = addDataRow_(REPORT, row, 'SMA50', lookup('N'), '$#,##0.00');
  row = addDataRow_(REPORT, row, 'SMA200', lookup('O'), '$#,##0.00');
  row = addDataRow_(REPORT, row, 'ADX', lookup('S'), '0.00');
  row = addDataRow_(REPORT, row, 'ATR', lookup('X'), '0.00');
  
  // MOMENTUM & TIMING Section
  row = addSection_(REPORT, row, 'MOMENTUM & TIMING');
  row = addDataRow_(REPORT, row, 'RSI', lookup('P'), '0.00');
  row = addDataRow_(REPORT, row, 'MACD Hist', lookup('Q'), '0.000');
  row = addDataRow_(REPORT, row, 'Divergence', lookup('R'), '@');
  row = addDataRow_(REPORT, row, 'Stoch %K', lookup('T'), '0.0%');
  row = addDataRow_(REPORT, row, 'Bollinger %B', lookup('Y'), '0.0%');
  
  // LEVELS & PLANNING Section
  row = addSection_(REPORT, row, 'LEVELS & PLANNING');
  row = addDataRow_(REPORT, row, 'Support', lookup('U'), '$#,##0.00');
  row = addDataRow_(REPORT, row, 'Resistance', lookup('V'), '$#,##0.00');
  row = addDataRow_(REPORT, row, 'Target', lookup('W'), '$#,##0.00');
  row = addDataRow_(REPORT, row, 'ATH', lookup('H'), '$#,##0.00');
  row = addDataRow_(REPORT, row, 'ATH Diff', lookup('I'), '0.00%');
  
  // Narrative sections - only FUND NOTES
  row = addNarrative_(REPORT, row, 'FUND NOTES', `=IFERROR(INDEX(CALCULATIONS!AA:AA,MATCH(UPPER(TRIM($A$1)),ARRAYFORMULA(UPPER(TRIM(CALCULATIONS!A:A))),0)),"—")`);
  
  // Final styling
  const finalRow = REPORT.getLastRow();
  REPORT.getRange(1, 1, finalRow, 3)
    .setBorder(true, true, true, true, true, true, P.GRID, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  REPORT.setHiddenGridlines(true);
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
    .setFontWeight('bold')
    .setFontSize(11)
    .setFontFamily('Calibri')
    .setBorder(true, true, true, true, false, false, P.GRID, SpreadsheetApp.BorderStyle.SOLID)
    .setHorizontalAlignment('left');
  REPORT.setRowHeight(row, 22);
  return row + 1;
}

/**
 * Add data row with column C narrative and conditional formatting
 */
function addDataRow_(REPORT, row, label, formula, format) {
  const P = reportPalette___();
  
  // Label
  REPORT.getRange(row, 1).setValue(label);
  
  // Formula in column B
  REPORT.getRange(row, 2).setFormula(formula);
  
  // Add narrative in column C for split zone (rows 9-34)
  const isSplit = (row >= 9 && row <= 34);
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
  
  REPORT.getRange(row, 1).setFontColor(P.MUTED).setFontWeight('bold').setHorizontalAlignment('left').setFontFamily('Calibri');
  REPORT.getRange(row, 2).setFontColor(P.TEXT).setFontWeight('normal').setHorizontalAlignment('left').setFontFamily('Calibri');
  
  if (isSplit) {
    REPORT.getRange(row, 3).setFontColor(P.TEXT).setFontWeight('normal').setHorizontalAlignment('left').setWrap(true).setFontFamily('Calibri');
    REPORT.setRowHeight(row, 34); // Taller for narrative
  } else {
    REPORT.getRange(row, 2, 1, 2).setFontColor(P.TEXT).setFontWeight('normal').setHorizontalAlignment('left').setFontFamily('Calibri');
    REPORT.setRowHeight(row, 18);
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
    
    case 'RVOL':
      return '=IFERROR(IF(' + lookup('G') + '>=1.5,"strong participation.",IF(' + lookup('G') + '>=1,"average participation.","low participation (drift/chop risk).")),"—")';
    
    case 'TREND':
      return '=IFERROR(' + lookup('L') + ' & IF(ISBLANK(' + lookup('K') + '),"",", Score: " & ' + lookup('K') + '),"—")';
    
    case 'R:R':
      return '=IFERROR(IF(' + lookup('J') + '>=3,"elite asymmetry.",IF(' + lookup('J') + '>=1.5,"acceptable asymmetry.","weak asymmetry.")),"—")';
    
    case 'P/E':
      return '=IFERROR("P/E ratio: " & TEXT(GOOGLEFINANCE($A$1,"pe"),"0.00"),"P/E data unavailable")';
    
    case 'EPS':
      return '=IFERROR("Earnings per share: $" & TEXT(GOOGLEFINANCE($A$1,"eps"),"0.00"),"EPS data unavailable")';
    
    case 'SMA20':
      return '=IFERROR(TEXT((' + numLookup('E') + '/' + numLookup('M') + '-1),"+0.0%;-0.0%") & " (price " & TEXT(' + numLookup('E') + ',\"0.00\") & IF(' + numLookup('E') + '>=' + numLookup('M') + '," > SMA20)."," < SMA20)."),\"—\")';
    
    case 'SMA50':
      return '=IFERROR(TEXT((' + numLookup('E') + '/' + numLookup('N') + '-1),"+0.0%;-0.0%") & " (price " & TEXT(' + numLookup('E') + ',\"0.00\") & IF(' + numLookup('E') + '>=' + numLookup('N') + '," > SMA50)."," < SMA50)."),\"—\")';
    
    case 'SMA200':
      return '=IFERROR(TEXT((' + numLookup('E') + '/' + numLookup('O') + '-1),"+0.0%;-0.0%") & " (price " & TEXT(' + numLookup('E') + ',\"0.00\") & IF(' + numLookup('E') + '>=' + numLookup('O') + '," > SMA200). Risk-On."," < SMA200). Risk-Off."),\"—\")';
    
    case 'ADX':
      return '=IFERROR(IF(' + lookup('S') + '>=25,"strong trend.",IF(' + lookup('S') + '>=20,"trend developing.",IF(' + lookup('S') + '>=15,"weak trend.","range-bound."))),"—")';
    
    case 'ATR':
      return '=IFERROR("Average True Range: " & TEXT(' + lookup('X') + ',"$0.00") & " (" & TEXT(' + lookup('X') + '/' + lookup('E') + ',"0.0%") & " of price).","—")';
    
    case 'RSI':
      return '=IFERROR(IF(' + lookup('P') + '>=70,"overbought.",IF(' + lookup('P') + '<=30,"oversold.",IF(' + lookup('P') + '>=55,"positive momentum.",IF(' + lookup('P') + '<=45,"weak momentum.","neutral.")))),"—")';
    
    case 'MACD Hist':
      return '=IFERROR(IF(' + lookup('Q') + '>0,"positive impulse.",IF(' + lookup('Q') + '<0,"negative impulse.","flat impulse.")),"—")';
    
    case 'Divergence':
      return '=IFERROR(IF(ISBLANK(' + lookup('R') + '),"No divergence detected.",' + lookup('R') + '),"—")';
    
    case 'Stoch %K':
      return '=IFERROR(IF(' + lookup('T') + '>=0.8,"overbought timing.",IF(' + lookup('T') + '<=0.2,"oversold timing.","neutral timing.")),"—")';
    
    case 'Bollinger %B':
      return '=IFERROR(IF(' + lookup('Y') + '>1,"above upper band (expansion).",IF(' + lookup('Y') + '>=0.8,"upper-band zone.",IF(' + lookup('Y') + '<0,"below lower band (extreme).",IF(' + lookup('Y') + '<=0.2,"lower-band zone.","mid-band zone.")))),"—")';
    
    case 'Support':
      return '=IFERROR("Price " & TEXT(' + numLookup('E') + ',\"0.00\") & " " & TEXT((' + numLookup('E') + '/' + numLookup('U') + '-1),"+0.0%;-0.0%") & IF(' + numLookup('E') + '<' + numLookup('U') + '," below support."," above support."),\"—\")';
    
    case 'Resistance':
      return '=IFERROR("Price " & TEXT(' + numLookup('E') + ',\"0.00\") & " " & IF(' + numLookup('E') + '>' + numLookup('V') + ',TEXT((1-' + numLookup('E') + '/' + numLookup('V') + '),"+0.0%;-0.0%") & " above resistance (R:" & TEXT(' + numLookup('V') + ',\"0.00\") & ").",TEXT((' + numLookup('V') + '/' + numLookup('E') + '-1),"+0.0%;-0.0%") & " below resistance (R:" & TEXT(' + numLookup('V') + ',\"0.00\") & ")."),\"—\")';
    
    case 'Target':
      return '=IFERROR("Target " & TEXT(' + numLookup('W') + ',"$#,##0.00") & " (" & TEXT((' + numLookup('W') + '/' + numLookup('E') + '-1),"+0.00%;-0.00%") & ").","—")';
    
    case 'ATH':
      return '=IFERROR("ATH " & TEXT(' + numLookup('H') + ',"$#,##0.00") & " (" & TEXT((' + numLookup('E') + '/' + numLookup('H') + '-1),"+0.00%;-0.00%") & ").","—")';
    
    case 'ATH Diff':
      return '=IFERROR(TEXT(' + lookup('I') + ',"+0.00%;-0.00%") & " vs ATH.","—")';
    
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
      
    case 'R:R':
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
      
    case 'ATH DIFF':
      rules.push(
        SpreadsheetApp.newConditionalFormatRule()
          .whenNumberGreaterThanOrEqualTo(0)
          .setBackground(P.CHIP_POS)
          .setFontColor(P.POS_TXT)
          .setRanges([valueCell])
          .build(),
        SpreadsheetApp.newConditionalFormatRule()
          .whenNumberBetween(-0.05, -0.001)
          .setBackground(P.CHIP_WARN)
          .setFontColor(P.WARN_TXT)
          .setRanges([valueCell])
          .build(),
        SpreadsheetApp.newConditionalFormatRule()
          .whenNumberLessThan(-0.05)
          .setBackground(P.CHIP_NEG)
          .setFontColor(P.NEG_TXT)
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
    .setFontWeight('bold')
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
 * Setup ticker dropdown
 */
function setupReportTickerDropdown_(reportSheet, inputSheet) {
  const last = inputSheet.getLastRow();
  const height = Math.max(1, last - 2);
  const rng = inputSheet.getRange(3, 1, height, 1);

  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(rng, true)
    .setAllowInvalid(false)
    .build();

  const a1 = reportSheet.getRange('A1');
  a1.setDataValidation(rule);
  a1.setFontWeight('bold');
  a1.setFontFamily('Calibri');
  a1.setHorizontalAlignment('left');
}

/**
 * Set column widths
 */
function setReportColumnWidthsAndWrap___(REPORT) {
  const pxPerChar = 8;
  const colA = Math.max(70, Math.round(8 * pxPerChar + 15));
  const colB = Math.max(70, Math.round(8 * pxPerChar + 15));
  const colC = Math.max(170, Math.round(27 * pxPerChar + 20));

  REPORT.setColumnWidth(1, colA);
  REPORT.setColumnWidth(2, colB);
  REPORT.setColumnWidth(3, colC);

  const lastRow = Math.max(1, REPORT.getLastRow());
  REPORT.getRange(1, 1, lastRow, 3).setWrap(true);
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

 * Apply conditional formatting to decision cells (B2:B4)
 */
function applyDecisionConditionalFormatting_(REPORT) {
  const P = reportPalette___();
  
  // SIGNAL cell (B2)
  const signalCell = REPORT.getRange('B2');
  const signalRules = [
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains('BUY')
      .setBackground(P.CHIP_POS)
      .setFontColor(P.POS_TXT)
      .setRanges([signalCell])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains('BULL')
      .setBackground(P.CHIP_POS)
      .setFontColor(P.POS_TXT)
      .setRanges([signalCell])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains('RISK-ON')
      .setBackground(P.CHIP_POS)
      .setFontColor(P.POS_TXT)
      .setRanges([signalCell])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains('Breakout')
      .setBackground(P.CHIP_POS)
      .setFontColor(P.POS_TXT)
      .setRanges([signalCell])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains('Trend Continuation')
      .setBackground(P.CHIP_POS)
      .setFontColor(P.POS_TXT)
      .setRanges([signalCell])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains('SELL')
      .setBackground(P.CHIP_NEG)
      .setFontColor(P.NEG_TXT)
      .setRanges([signalCell])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains('BEAR')
      .setBackground(P.CHIP_NEG)
      .setFontColor(P.NEG_TXT)
      .setRanges([signalCell])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains('RISK-OFF')
      .setBackground(P.CHIP_NEG)
      .setFontColor(P.NEG_TXT)
      .setRanges([signalCell])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains('Stop-Out')
      .setBackground(P.CHIP_NEG)
      .setFontColor(P.NEG_TXT)
      .setRanges([signalCell])
      .build()
  ];
  
  // FUNDAMENTAL cell (B3)
  const fundamentalCell = REPORT.getRange('B3');
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
  
  // DECISION cell (B4)
  const decisionCell = REPORT.getRange('B4');
  const decisionRules = [
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
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains('HOLD')
      .setBackground(P.CHIP_NEU)
      .setFontColor(P.TEXT)
      .setRanges([decisionCell])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains('Take Profit')
      .setBackground(P.CHIP_WARN)
      .setFontColor(P.WARN_TXT)
      .setRanges([decisionCell])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains('Reduce')
      .setBackground(P.CHIP_WARN)
      .setFontColor(P.WARN_TXT)
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