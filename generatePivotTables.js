/**
 * ==============================================================================
 * INSTITUTIONAL-GRADE PIVOT TABLES FOR QUANTITATIVE STOCK ANALYSIS
 * ==============================================================================
 */

function generatePivotTablesSheet() {
  const startTime = new Date();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  try {
    const calcSheet = ss.getSheetByName("CALCULATIONS");
    const dataSheet = ss.getSheetByName("DATA");
    
    if (!calcSheet) {
      ss.toast('CALCULATIONS sheet not found. Generate it first.', '❌ Error', 3);
      return;
    }

    ss.toast('Building institutional-grade pivot analysis...', '⏳ Processing', 2);
    
    let pivotSheet = ss.getSheetByName("PIVOT_ANALYSIS") || ss.insertSheet("PIVOT_ANALYSIS");
    pivotSheet.clear().clearFormats();
    
    const SEP = (/^(en|en_)/.test(ss.getSpreadsheetLocale())) ? "," : ";";
    const marketRegime = dataSheet ? getMarketRegime(dataSheet) : { usa: "NEUTRAL", india: "NEUTRAL" };
    
    buildInstitutionalPivotLayout(pivotSheet, SEP, marketRegime);
    
    const elapsed = ((new Date() - startTime) / 1000).toFixed(2);
    ss.toast(`✓ Institutional pivot analysis generated in ${elapsed}s`, 'Success', 3);
    
  } catch (error) {
    Logger.log(`Error: ${error.stack}`);
    ss.toast(`Failed: ${error.message}`, '❌ Error', 5);
  }
}

function getMarketRegime(dataSheet) {
  try {
    const row3 = dataSheet.getRange("3:3").getValues()[0];
    let usaRegimeValue = "NEUTRAL";
    let indiaRegimeValue = "NEUTRAL";
    
    for (let i = 0; i < row3.length; i++) {
      if (row3[i] && typeof row3[i] === 'string') {
        const val = String(row3[i]).toUpperCase();
        if (val.includes("BULL") || val.includes("BEAR")) {
          if (usaRegimeValue === "NEUTRAL") {
            usaRegimeValue = val;
          } else if (indiaRegimeValue === "NEUTRAL") {
            indiaRegimeValue = val;
            break;
          }
        }
      }
    }
    
    return { usa: usaRegimeValue, india: indiaRegimeValue };
  } catch (error) {
    return { usa: "NEUTRAL", india: "NEUTRAL" };
  }
}

function buildInstitutionalPivotLayout(sheet, SEP, marketRegime) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  sheet.getRange("A1:M1").merge()
    .setValue("📊 INSTITUTIONAL QUANTITATIVE FACTOR ANALYSIS")
    .setBackground("#0D47A1")
    .setFontColor("#FFFFFF")
    .setFontWeight("bold")
    .setFontSize(16)
    .setHorizontalAlignment("center");
  
  const timestamp = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "yyyy-MM-dd HH:mm:ss");
  sheet.getRange("A2")
    .setValue(`Market: USA ${marketRegime.usa} | INDIA ${marketRegime.india} | ${timestamp}`)
    .setFontSize(9)
    .setFontColor("#1565C0");
  sheet.getRange("A2:M2").merge();
  
  let currentRow = 4;
  currentRow = buildMultiFactorScoring(sheet, currentRow, SEP);
  currentRow += 2;
  currentRow = buildRegimeAnalysis(sheet, currentRow, SEP);
  currentRow += 2;
  currentRow = buildRiskMetrics(sheet, currentRow, SEP);
  currentRow += 2;
  currentRow = buildTopOpportunities(sheet, currentRow, SEP);
  
  sheet.setColumnWidths(1, 13, 110);
  sheet.setColumnWidth(1, 180);
}

function buildMultiFactorScoring(sheet, startRow, SEP) {
  sheet.getRange(startRow, 1, 1, 8).merge()
    .setValue("1. MULTI-FACTOR SCORING & PERFORMANCE")
    .setBackground("#1565C0")
    .setFontColor("#FFFFFF")
    .setFontWeight("bold")
    .setFontSize(11);
  startRow++;
  
  const headers = ["Factor", "Weight", "Avg Score", "Top Q", "Bottom Q", "Spread", "Sharpe", "Hit Rate"];
  sheet.getRange(startRow, 1, 1, 8).setValues([headers])
    .setBackground("#E3F2FD")
    .setFontWeight("bold")
    .setHorizontalAlignment("center");
  startRow++;
  
  const factors = [
    ["Momentum", "25%", 
     `=AVERAGE(CALCULATIONS!R:R)/100`,
     `=PERCENTILE(CALCULATIONS!H:H${SEP}0.75)`,
     `=PERCENTILE(CALCULATIONS!H:H${SEP}0.25)`,
     `=D${startRow}-E${startRow}`,
     `=AVERAGE(CALCULATIONS!H:H)/STDEV(CALCULATIONS!H:H)`,
     `=COUNTIF(CALCULATIONS!H:H${SEP}">0")/COUNTA(CALCULATIONS!H:H)`],
    ["Trend", "20%",
     `=COUNTIF(CALCULATIONS!N:N${SEP}"BULL")/COUNTA(CALCULATIONS!N:N)`,
     `=AVERAGEIFS(CALCULATIONS!H:H${SEP}CALCULATIONS!N:N${SEP}"BULL")`,
     `=AVERAGEIFS(CALCULATIONS!H:H${SEP}CALCULATIONS!N:N${SEP}"BEAR")`,
     `=D${startRow+1}-E${startRow+1}`,
     `=D${startRow+1}/STDEV(CALCULATIONS!H:H)`,
     `=COUNTIFS(CALCULATIONS!N:N${SEP}"BULL"${SEP}CALCULATIONS!H:H${SEP}">0")/COUNTIF(CALCULATIONS!N:N${SEP}"BULL")`],
    ["Value", "15%",
     `=AVERAGE(CALCULATIONS!K:K)`,
     `=AVERAGEIFS(CALCULATIONS!H:H${SEP}CALCULATIONS!K:K${SEP}"<-0.3")`,
     `=AVERAGEIFS(CALCULATIONS!H:H${SEP}CALCULATIONS!K:K${SEP}">-0.05")`,
     `=D${startRow+2}-E${startRow+2}`,
     `=D${startRow+2}/STDEV(CALCULATIONS!H:H)`,
     `=COUNTIFS(CALCULATIONS!K:K${SEP}"<-0.2"${SEP}CALCULATIONS!H:H${SEP}">0")/COUNTIFS(CALCULATIONS!K:K${SEP}"<-0.2")`]
  ];
  
  sheet.getRange(startRow, 1, 3, 8).setFormulas(factors);
  sheet.getRange(startRow, 2, 3, 1).setNumberFormat("0%");
  sheet.getRange(startRow, 3, 3, 5).setNumberFormat("0.00");
  sheet.getRange(startRow, 8, 3, 1).setNumberFormat("0.0%");
  
  return startRow + 3;
}

function buildRegimeAnalysis(sheet, startRow, SEP) {
  sheet.getRange(startRow, 1, 1, 8).merge()
    .setValue("2. REGIME-BASED PERFORMANCE")
    .setBackground("#00838F")
    .setFontColor("#FFFFFF")
    .setFontWeight("bold")
    .setFontSize(11);
  startRow++;
  
  const headers = ["Regime", "Count", "Avg Return", "Win Rate", "Best", "Worst", "Sharpe", "Sortino"];
  sheet.getRange(startRow, 1, 1, 8).setValues([headers])
    .setBackground("#E0F2F1")
    .setFontWeight("bold")
    .setHorizontalAlignment("center");
  startRow++;
  
  const regimes = [
    ["BULL",
     `=COUNTIF(CALCULATIONS!N:N${SEP}"BULL")`,
     `=AVERAGEIF(CALCULATIONS!N:N${SEP}"BULL"${SEP}CALCULATIONS!H:H)`,
     `=COUNTIFS(CALCULATIONS!N:N${SEP}"BULL"${SEP}CALCULATIONS!H:H${SEP}">0")/B${startRow}`,
     `=MAXIFS(CALCULATIONS!H:H${SEP}CALCULATIONS!N:N${SEP}"BULL")`,
     `=MINIFS(CALCULATIONS!H:H${SEP}CALCULATIONS!N:N${SEP}"BULL")`,
     `=C${startRow}/STDEV(CALCULATIONS!H:H)`,
     `=C${startRow}/STDEV(CALCULATIONS!H:H)*1.5`],
    ["BEAR",
     `=COUNTIF(CALCULATIONS!N:N${SEP}"BEAR")`,
     `=AVERAGEIF(CALCULATIONS!N:N${SEP}"BEAR"${SEP}CALCULATIONS!H:H)`,
     `=COUNTIFS(CALCULATIONS!N:N${SEP}"BEAR"${SEP}CALCULATIONS!H:H${SEP}">0")/B${startRow+1}`,
     `=MAXIFS(CALCULATIONS!H:H${SEP}CALCULATIONS!N:N${SEP}"BEAR")`,
     `=MINIFS(CALCULATIONS!H:H${SEP}CALCULATIONS!N:N${SEP}"BEAR")`,
     `=C${startRow+1}/STDEV(CALCULATIONS!H:H)`,
     `=C${startRow+1}/STDEV(CALCULATIONS!H:H)*1.5`],
    ["HIGH VOL",
     `=COUNTIF(CALCULATIONS!W:W${SEP}"HIGH VOL")`,
     `=AVERAGEIF(CALCULATIONS!W:W${SEP}"HIGH VOL"${SEP}CALCULATIONS!H:H)`,
     `=COUNTIFS(CALCULATIONS!W:W${SEP}"HIGH VOL"${SEP}CALCULATIONS!H:H${SEP}">0")/B${startRow+2}`,
     `=MAXIFS(CALCULATIONS!H:H${SEP}CALCULATIONS!W:W${SEP}"HIGH VOL")`,
     `=MINIFS(CALCULATIONS!H:H${SEP}CALCULATIONS!W:W${SEP}"HIGH VOL")`,
     `=C${startRow+2}/STDEV(CALCULATIONS!H:H)`,
     `=C${startRow+2}/STDEV(CALCULATIONS!H:H)*1.5`]
  ];
  
  sheet.getRange(startRow, 1, 3, 8).setFormulas(regimes);
  sheet.getRange(startRow, 3, 3, 3).setNumberFormat("0.00%");
  sheet.getRange(startRow, 4, 3, 1).setNumberFormat("0.0%");
  sheet.getRange(startRow, 7, 3, 2).setNumberFormat("0.00");
  
  return startRow + 3;
}

function buildRiskMetrics(sheet, startRow, SEP) {
  sheet.getRange(startRow, 1, 1, 6).merge()
    .setValue("3. RISK-ADJUSTED METRICS")
    .setBackground("#AD1457")
    .setFontColor("#FFFFFF")
    .setFontWeight("bold")
    .setFontSize(11);
  startRow++;
  
  const headers = ["Metric", "Value", "Benchmark", "Status", "Percentile", "Quality"];
  sheet.getRange(startRow, 1, 1, 6).setValues([headers])
    .setBackground("#F8BBD0")
    .setFontWeight("bold")
    .setHorizontalAlignment("center");
  startRow++;
  
  const metrics = [
    ["Sharpe Ratio",
     `=AVERAGE(CALCULATIONS!H:H)/STDEV(CALCULATIONS!H:H)`,
     "1.00",
     `=IF(B${startRow}>1.5${SEP}"Excellent"${SEP}IF(B${startRow}>1${SEP}"Good"${SEP}"Fair"))`,
     `=PERCENTRANK(CALCULATIONS!H:H${SEP}AVERAGE(CALCULATIONS!H:H))`,
     `=IF(B${startRow}>1.5${SEP}"HIGH"${SEP}IF(B${startRow}>1${SEP}"MEDIUM"${SEP}"LOW"))`],
    ["Win Rate",
     `=COUNTIF(CALCULATIONS!H:H${SEP}">0")/COUNTA(CALCULATIONS!H:H)`,
     "0.50",
     `=IF(B${startRow+1}>0.6${SEP}"High"${SEP}IF(B${startRow+1}>0.5${SEP}"Moderate"${SEP}"Low"))`,
     `=B${startRow+1}`,
     `=IF(B${startRow+1}>0.6${SEP}"HIGH"${SEP}IF(B${startRow+1}>0.5${SEP}"MEDIUM"${SEP}"LOW"))`],
    ["Profit Factor",
     `=SUMIF(CALCULATIONS!H:H${SEP}">0")/ABS(SUMIF(CALCULATIONS!H:H${SEP}"<0"))`,
     "1.50",
     `=IF(B${startRow+2}>2${SEP}"Excellent"${SEP}IF(B${startRow+2}>1.5${SEP}"Good"${SEP}"Fair"))`,
     `=PERCENTRANK(CALCULATIONS!H:H${SEP}AVERAGE(CALCULATIONS!H:H))`,
     `=IF(B${startRow+2}>2${SEP}"HIGH"${SEP}IF(B${startRow+2}>1.5${SEP}"MEDIUM"${SEP}"LOW"))`]
  ];
  
  sheet.getRange(startRow, 1, 3, 6).setFormulas(metrics);
  sheet.getRange(startRow, 2, 3, 1).setNumberFormat("0.00");
  sheet.getRange(startRow, 3, 3, 1).setNumberFormat("0.00");
  sheet.getRange(startRow, 5, 3, 1).setNumberFormat("0.0%");
  
  return startRow + 3;
}

function buildTopOpportunities(sheet, startRow, SEP) {
  sheet.getRange(startRow, 1, 1, 8).merge()
    .setValue("4. TOP OPPORTUNITIES (Composite Score)")
    .setBackground("#2E7D32")
    .setFontColor("#FFFFFF")
    .setFontWeight("bold")
    .setFontSize(11);
  startRow++;
  
  const headers = ["Ticker", "Score", "Decision", "Signal", "R:R", "Change%", "ATH Zone", "Pattern"];
  sheet.getRange(startRow, 1, 1, 8).setValues([headers])
    .setBackground("#C8E6C9")
    .setFontWeight("bold")
    .setHorizontalAlignment("center");
  startRow++;
  
  const formula = 
    `=QUERY({CALCULATIONS!A:A${SEP}` +
    `ARRAYFORMULA((CALCULATIONS!R:R/100)*0.3+(IF(CALCULATIONS!N:N="BULL"${SEP}1${SEP}0))*0.25+(ABS(CALCULATIONS!K:K))*0.20+(IF(CALCULATIONS!W:W="HIGH VOL"${SEP}1${SEP}0))*0.15+(IF(CALCULATIONS!M:M="VALUE"${SEP}1${SEP}0))*0.10)${SEP}` +
    `CALCULATIONS!C:C${SEP}CALCULATIONS!D:D${SEP}CALCULATIONS!AB:AB${SEP}CALCULATIONS!H:H${SEP}CALCULATIONS!L:L${SEP}CALCULATIONS!E:E}${SEP}` +
    `"SELECT Col1${SEP}Col2${SEP}Col3${SEP}Col4${SEP}Col5${SEP}Col6${SEP}Col7${SEP}Col8 ` +
    `WHERE Col1<>'' AND Col1<>'Ticker' AND (Col3 CONTAINS 'BUY' OR Col3 CONTAINS 'TRADE') AND Col5>=2 ` +
    `ORDER BY Col2 DESC LIMIT 15")`;
  
  sheet.getRange(startRow, 1).setFormula(formula);
  sheet.getRange(startRow, 2, 15, 1).setNumberFormat("0.000");
  sheet.getRange(startRow, 5, 15, 1).setNumberFormat("0.00");
  sheet.getRange(startRow, 6, 15, 1).setNumberFormat("0.00%");
  
  return startRow + 15;
}
