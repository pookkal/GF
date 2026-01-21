/**
 * ==============================================================================
 * BUY CATEGORIES DASHBOARD - Single Filtered Table
 * One master table with dropdown filters for different buy strategies
 * ==============================================================================
 */

function generateBuyCategoriesSheet() {
  const startTime = new Date();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  try {
    const calcSheet = ss.getSheetByName("CALCULATIONS");
    if (!calcSheet) {
      ss.toast('CALCULATIONS sheet not found. Generate it first.', '❌ Error', 3);
      return;
    }

    const lastRow = calcSheet.getLastRow();
    if (lastRow < 3) {
      ss.toast('CALCULATIONS sheet has no data rows. Generate it first.', '❌ Error', 3);
      return;
    }

    ss.toast('Building Buy Categories Dashboard...', '⏳ Processing', 2);
    
    let buySheet = ss.getSheetByName("BUY_CATEGORIES") || ss.insertSheet("BUY_CATEGORIES");
    buySheet.clear().clearFormats();
    
    const SEP = (/^(en|en_)/.test(ss.getSpreadsheetLocale())) ? "," : ";";
    
    buildSingleTableWithFilters(buySheet, SEP, ss);
    
    const elapsed = ((new Date() - startTime) / 1000).toFixed(2);
    ss.toast(`✓ Buy Categories Dashboard generated in ${elapsed}s`, 'Success', 3);
    
  } catch (error) {
    Logger.log(`Error: ${error.stack}`);
    ss.toast(`Failed: ${error.message}`, '❌ Error', 5);
  }
}

function buildSingleTableWithFilters(sheet, SEP, ss) {
  // Title
  sheet.getRange("A1:L1").merge()
    .setValue("🎯 BUY CATEGORIES - Filtered Stock Selection")
    .setBackground("#1A237E")
    .setFontColor("#FFFFFF")
    .setFontWeight("bold")
    .setFontSize(14)
    .setHorizontalAlignment("center");
  
  // Filter Controls Row
  sheet.getRange("A2").setValue("📊 CATEGORY FILTER:").setFontWeight("bold").setBackground("#E8EAF6");
  sheet.getRange("B2").setValue("ALL BUY SIGNALS").setBackground("#FFF9C4");
  
  sheet.getRange("D2").setValue("MIN R:R:").setFontWeight("bold").setBackground("#E8EAF6");
  sheet.getRange("E2").setValue(1.5).setBackground("#FFF9C4").setNumberFormat("0.0");
  
  sheet.getRange("G2").setValue("MIN RSI:").setFontWeight("bold").setBackground("#E8EAF6");
  sheet.getRange("H2").setValue(30).setBackground("#FFF9C4").setNumberFormat("0");
  
  sheet.getRange("J2").setValue("MAX RSI:").setFontWeight("bold").setBackground("#E8EAF6");
  sheet.getRange("K2").setValue(70).setBackground("#FFF9C4").setNumberFormat("0");
  
  // Category dropdown in B2
  const categories = [
    "ALL BUY SIGNALS",
    "STRONG BUY (R:R >= 1.5)",
    "MOMENTUM BREAKOUT",
    "VALUE PLAYS",
    "OVERSOLD BOUNCE",
    "ACCUMULATION ZONE",
    "HIGH R:R SETUPS",
    "PATTERN CONFIRMED",
    "NEAR ATH"
  ];
  
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(categories, true)
    .build();
  sheet.getRange("B2").setDataValidation(rule);
  
  // Instructions
  sheet.getRange("A3:L3").merge()
    .setValue("� Select a category from dropdown above. Adjust R:R and RSI filters as needed. Table updates automatically.")
    .setBackground("#E3F2FD")
    .setFontSize(9)
    .setFontStyle("italic")
    .setHorizontalAlignment("center");
  
  // Column Headers
  const headers = ["Ticker", "Decision", "Signal", "Pattern", "Price", "Change%", "RSI", "R:R", "Trend", "Vol Regime", "ATH Diff%", "Target"];
  sheet.getRange(4, 1, 1, headers.length)
    .setValues([headers])
    .setBackground("#37474F")
    .setFontColor("#FFFFFF")
    .setFontWeight("bold")
    .setHorizontalAlignment("center");
  
  // Dynamic QUERY formula that responds to filters
  const formula = buildDynamicQueryFormula(SEP);
  sheet.getRange("A5").setFormula(formula);
  
  // Format data area
  const dataRange = sheet.getRange(5, 1, 50, headers.length);
  dataRange.setHorizontalAlignment("left")
    .setVerticalAlignment("middle")
    .setWrap(false)
    .setBorder(true, true, true, true, true, true, "#CCCCCC", SpreadsheetApp.BorderStyle.SOLID);
  
  // Format specific columns
  sheet.getRange(5, 6, 50, 1).setNumberFormat("0.00%");  // Change%
  sheet.getRange(5, 7, 50, 1).setNumberFormat("0.0");    // RSI
  sheet.getRange(5, 8, 50, 1).setNumberFormat("0.0");    // R:R
  sheet.getRange(5, 11, 50, 1).setNumberFormat("0.00%"); // ATH Diff%
  sheet.getRange(5, 5, 50, 1).setNumberFormat("#,##0.00"); // Price
  sheet.getRange(5, 12, 50, 1).setNumberFormat("#,##0.00"); // Target
  
  // Set column widths
  sheet.setColumnWidth(1, 90);   // Ticker
  sheet.setColumnWidth(2, 120);  // Decision
  sheet.setColumnWidth(3, 100);  // Signal
  sheet.setColumnWidth(4, 150);  // Pattern
  sheet.setColumnWidth(5, 80);   // Price
  sheet.setColumnWidth(6, 80);   // Change%
  sheet.setColumnWidth(7, 70);   // RSI
  sheet.setColumnWidth(8, 70);   // R:R
  sheet.setColumnWidth(9, 90);   // Trend
  sheet.setColumnWidth(10, 110); // Vol Regime
  sheet.setColumnWidth(11, 90);  // ATH Diff%
  sheet.setColumnWidth(12, 80);  // Target
  
  // Freeze header rows
  sheet.setFrozenRows(4);
  
  // Add category descriptions
  addCategoryLegend(sheet);
}

function buildDynamicQueryFormula(SEP) {
  // This formula reads the category from B2 and filters from E2, H2, K2
  // Then applies the appropriate filter to CALCULATIONS data
  
  const formula = `=IFERROR(
  QUERY(CALCULATIONS!A:AG${SEP}
    "SELECT A${SEP}C${SEP}D${SEP}E${SEP}G${SEP}H${SEP}R${SEP}AB${SEP}N${SEP}W${SEP}K${SEP}AA 
    WHERE A<>'' AND A<>'Ticker' 
    AND " & 
    IF($B$2="ALL BUY SIGNALS"${SEP}
      "(C CONTAINS 'BUY' OR C CONTAINS 'TRADE' OR C CONTAINS 'HOLD' OR D CONTAINS 'BUY')"${SEP}
    IF($B$2="STRONG BUY (R:R >= 1.5)"${SEP}
      "AB>=" & $E$2${SEP}
    IF($B$2="MOMENTUM BREAKOUT"${SEP}
      "I>1.0 AND N='BULL' AND R>45"${SEP}
    IF($B$2="VALUE PLAYS"${SEP}
      "K<-0.10 AND R<55"${SEP}
    IF($B$2="OVERSOLD BOUNCE"${SEP}
      "R<=" & $H$2${SEP}
    IF($B$2="ACCUMULATION ZONE"${SEP}
      "N='BULL' AND R>=" & $H$2 & " AND R<=" & $K$2${SEP}
    IF($B$2="HIGH R:R SETUPS"${SEP}
      "AB>=" & $E$2${SEP}
    IF($B$2="PATTERN CONFIRMED"${SEP}
      "E<>'' AND E<>'—'"${SEP}
    IF($B$2="NEAR ATH"${SEP}
      "K>=-0.15 AND N='BULL'"${SEP}
      "A<>''"
    ))))))))) &
    " AND R>=" & $H$2 & " AND R<=" & $K$2 &
    " ORDER BY H DESC LIMIT 50 
    LABEL A 'Ticker'${SEP}C 'Decision'${SEP}D 'Signal'${SEP}E 'Pattern'${SEP}G 'Price'${SEP}H 'Change%'${SEP}R 'RSI'${SEP}AB 'R:R'${SEP}N 'Trend'${SEP}W 'Vol Regime'${SEP}K 'ATH Diff%'${SEP}AA 'Target'"
  )${SEP}
  "No matches found - try adjusting filters"
)`;
  
  return formula;
}

function addCategoryLegend(sheet) {
  const startRow = 57;
  
  sheet.getRange(startRow, 1, 1, 12).merge()
    .setValue("📖 CATEGORY GUIDE")
    .setBackground("#37474F")
    .setFontColor("#FFFFFF")
    .setFontWeight("bold")
    .setHorizontalAlignment("center");
  
  const legend = [
    ["ALL BUY SIGNALS", "All stocks with any BUY, TRADE, or HOLD signal"],
    ["STRONG BUY", "High conviction plays with good risk/reward ratio"],
    ["MOMENTUM BREAKOUT", "Breaking out with volume, strong uptrend"],
    ["VALUE PLAYS", "Deep discount from ATH, not overbought"],
    ["OVERSOLD BOUNCE", "Oversold RSI, ready for mean reversion"],
    ["ACCUMULATION ZONE", "Bullish trend, RSI in neutral zone"],
    ["HIGH R:R SETUPS", "Best risk/reward opportunities"],
    ["PATTERN CONFIRMED", "Bullish chart patterns detected"],
    ["NEAR ATH", "Market leaders within 15% of all-time high"]
  ];
  
  sheet.getRange(startRow + 1, 1, legend.length, 2)
    .setValues(legend)
    .setBackground("#F5F5F5")
    .setBorder(true, true, true, true, true, true, "#CCCCCC", SpreadsheetApp.BorderStyle.SOLID);
  
  sheet.getRange(startRow + 1, 1, legend.length, 1).setFontWeight("bold");
}
