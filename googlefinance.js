/**
 * ==============================================================================
 * BASELINE LABEL: STABLE_MASTER_V55_DIVERGENCE_INDEX_FIX
 * DATE: 22 DEC 2025
 * FIX: Replaced volatile OFFSET with stable INDEX for Divergence calculations.
 * ==============================================================================
 */

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📈 Institutional Terminal')
    .addItem('🚀 1-CLICK REBUILD ALL', 'FlushAllSheetsAndBuild')
    .addSeparator()
    .addItem('1. Fetch Data Only', 'generateDataSheet')
    .addItem('2. Build Calculations', 'generateCalculationsSheet')
    .addItem('3. Refresh Dashboard Only', 'generateDashboardSheet')
    .addItem('4. Setup Chart Only', 'setupChartSheet')
    .addToUi();
}

function FlushAllSheetsAndBuild() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetsToDelete = ["DATA", "CALCULATIONS", "CHART", "DASHBOARD"];
  const ui = SpreadsheetApp.getUi();
  if (ui.alert('🚨 FINAL FIX', 'Apply Divergence Index Fix?', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;

  sheetsToDelete.forEach(name => {
    let sheet = ss.getSheetByName(name);
    if (sheet) ss.deleteSheet(sheet);
  });

  ui.showModelessDialog(HtmlService.createHtmlOutput("<b>1/4:</b> Syncing Global Data..."), "Status");
  generateDataSheet();
  SpreadsheetApp.flush();
  Utilities.sleep(2000); 

  ui.showModelessDialog(HtmlService.createHtmlOutput("<b>2/4:</b> Stabilizing Divergence..."), "Status");
  generateCalculationsSheet();
  SpreadsheetApp.flush();

  ui.showModelessDialog(HtmlService.createHtmlOutput("<b>3/4:</b> Updating Dashboard..."), "Status");
  generateDashboardSheet(); 
  SpreadsheetApp.flush();

  ui.showModelessDialog(HtmlService.createHtmlOutput("<b>4/4:</b> Finalizing Report..."), "Status");
  setupChartSheet();
  ui.alert('✅ Golden Baseline Active', 'Divergence logic fixed (No more dashes).', ui.ButtonSet.OK);
}

function onEdit(e) {
  const sheet = e.range.getSheet();
  if (sheet.getName() === "INPUT" && (e.range.getA1Notation() === "B1" || e.range.getA1Notation() === "C1")) {
    generateDashboardSheet();
  }
  if (sheet.getName() !== "CHART") return;
  const watchList = ["B1", "D2", "A3", "B3", "C3"];
  if (watchList.includes(e.range.getA1Notation())) {
    updateDynamicChart();
  }
}

/**
 * 1. DASHBOARD ENGINE
 */
function generateDashboardSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inputSheet = ss.getSheetByName("INPUT");
  if (!inputSheet) return;
  let dashboard = ss.getSheetByName("DASHBOARD") || ss.insertSheet("DASHBOARD", inputSheet.getIndex() + 1);
  dashboard.clear().clearFormats();
  const headers = [["Ticker", "Price", "Change %", "DECISION", "ATH (TRUE)", "ATH Diff %", "R:R Quality", "Trend Score", "Trend State", "SMA 20", "SMA 50", "SMA 200", "Vol Trend", "RSI", "Divergence", "Support", "Target (3:1)", "Resistance", "ATR (14)", "Bollinger %B", "REASONING"]];
  dashboard.getRange(2, 1, 1, 21).setValues(headers).setBackground("#212121").setFontColor("white").setFontWeight("bold");

  const formula = '=IFERROR(SORT(FILTER(CALCULATIONS!$A$3:$U, ISNUMBER(MATCH(CALCULATIONS!$A$3:$A, FILTER(INPUT!$A$3:$A, ' +
    '(IF(OR(INPUT!$B$1="", INPUT!$B$1="ALL"), 1, REGEXMATCH(INPUT!$B$3:$B, "(?i)\\b(" & SUBSTITUTE(SUBSTITUTE(INPUT!$B$1, ", ", "|"), ",", "|") & ")\\b"))) * ' +
    '(IF(OR(INPUT!$C$1="", INPUT!$C$1="ALL"), 1, REGEXMATCH(INPUT!$C$3:$C, "(?i)\\b(" & SUBSTITUTE(SUBSTITUTE(INPUT!$C$1, ", ", "|"), ",", "|") & ")\\b")))' +
    '), 0))), 3, FALSE), "No Matches Found")';
  
  dashboard.getRange("A3").setFormula(formula);
  dashboard.getRange("A:S").setHorizontalAlignment("left");
  
  SpreadsheetApp.flush();
  dashboard.getRangeList(['C3:C', 'F3:F', 'T3:T']).setNumberFormat("0.00%");
  
  dashboard.setFrozenRows(2);
}

/**
 * 2. CHART SIDEBAR
 */
function setupChartSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inputSheet = ss.getSheetByName("INPUT");
  const tickers = getCleanTickers(inputSheet);
  let chartSheet = ss.getSheetByName("CHART") || ss.insertSheet("CHART");
  
  chartSheet.clear().clearFormats();
  forceExpandSheet(chartSheet, 45);
  
  chartSheet.setColumnWidth(1, 180); 
  chartSheet.setColumnWidth(2, 120); 
  chartSheet.setColumnWidth(5, 500); 

  chartSheet.getRange("A1").setValue("TICKER:").setFontWeight("bold");
  chartSheet.getRange("B1").setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(tickers).build()).setValue(tickers[0]);
  chartSheet.getRange("D1").setValue("VIEW:").setFontWeight("bold");
  chartSheet.getRange("D2").setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(["DAILY", "WEEKLY"]).build()).setValue("DAILY");
  
  chartSheet.getRange("E1").setFormula('=IFERROR(VLOOKUP(B1, CALCULATIONS!$A$3:$U, 21, 0), "—")').setWrap(true).setVerticalAlignment("top");

  chartSheet.getRange("A2:C2").setValues([["YEAR", "MONTH", "DAY"]]).setFontWeight("bold").setHorizontalAlignment("center");
  
  const numRule = (max) => SpreadsheetApp.newDataValidation().requireValueInList(Array.from({length: max + 1}, (_, i) => i)).build();
  chartSheet.getRange("A3").setDataValidation(numRule(5)).setValue(0);
  chartSheet.getRange("B3").setDataValidation(numRule(12)).setValue(3);
  chartSheet.getRange("C3").setDataValidation(numRule(31)).setValue(0);

  chartSheet.getRange("A4").setValue("DATE").setFontWeight("bold");
  chartSheet.getRange("B4").setFormula("=DATE(YEAR(TODAY())-A3, MONTH(TODAY())-B3, DAY(TODAY())-C3)").setNumberFormat("yyyy-mm-dd");

  chartSheet.getRange("A1:F4").setBackground("#000000").setFontColor("#FFFF00");
  chartSheet.getRangeList(["A3", "B3", "C3", "D2"]).setBackground("#FF80AB").setFontColor("#000000").setHorizontalAlignment("center").setFontWeight("bold");
  chartSheet.getRange("B1").setBackground("#FFEB3B").setFontColor("black").setFontWeight("bold");

  const t = "B1";
  
  const data = [
    ["SIGNAL (DECISION)", `=IFERROR(VLOOKUP(${t}, CALCULATIONS!$A$3:$U, 4, 0), "Wait")`], 
    ["LIVE PRICE", `=IFERROR(GOOGLEFINANCE(${t}, "price"), 0)`], 
    ["CHANGE ($)", `=IFERROR(B6 - GOOGLEFINANCE(${t}, "closeyest"), 0)`], 
    ["CHANGE (%)", `=IFERROR(GOOGLEFINANCE(${t}, "changepct")/100, 0)`], 
    ["", ""], 
    ["[ VALUATION METRICS ]", ""], 
    ["ATH (TRUE)", `=IFERROR(VLOOKUP(${t}, CALCULATIONS!$A$3:$U, 5, 0), 0)`], 
    ["DIFF FROM ATH %", `=IFERROR(VLOOKUP(${t}, CALCULATIONS!$A$3:$U, 6, 0), 0)`], 
    ["P/E RATIO", `=IFERROR(GOOGLEFINANCE(${t}, "pe"), 0)`],
    ["EPS", `=IFERROR(GOOGLEFINANCE(${t}, "eps"), 0)`],
    ["52W HIGH", `=IFERROR(GOOGLEFINANCE(${t}, "high52"), 0)`],
    ["52W LOW", `=IFERROR(GOOGLEFINANCE(${t}, "low52"), 0)`],
    ["", ""], 
    ["[ MOMENTUM & TREND ]", ""], 
    ["SMA 20", `=IFERROR(VLOOKUP(${t}, CALCULATIONS!$A$3:$U, 10, 0), 0)`], 
    ["SMA 50", `=IFERROR(VLOOKUP(${t}, CALCULATIONS!$A$3:$U, 11, 0), 0)`], 
    ["SMA 200", `=IFERROR(VLOOKUP(${t}, CALCULATIONS!$A$3:$U, 12, 0), 0)`], 
    ["RSI (14)", `=IFERROR(VLOOKUP(${t}, CALCULATIONS!$A$3:$U, 14, 0), 50)`], 
    ["TREND STATE", `=IFERROR(VLOOKUP(${t}, CALCULATIONS!$A$3:$U, 9, 0), "—")`], 
    ["DIVERGENCE", `=IFERROR(VLOOKUP(${t}, CALCULATIONS!$A$3:$U, 15, 0), "Neutral")`], 
    ["RELATIVE VOLUME", `=IFERROR(VLOOKUP(${t}, CALCULATIONS!$A$3:$U, 13, 0), 1)`], 
    ["", ""], 
    ["[ TECHNICAL LEVELS ]", ""], 
    ["SUPPORT FLOOR", `=IFERROR(VLOOKUP(${t}, CALCULATIONS!$A$3:$U, 16, 0), 0)`], 
    ["RESISTANCE CEILING", `=IFERROR(VLOOKUP(${t}, CALCULATIONS!$A$3:$U, 18, 0), 0)`], 
    ["TARGET (3:1 R:R)", `=IFERROR(VLOOKUP(${t}, CALCULATIONS!$A$3:$U, 17, 0), 0)`], 
    ["ATR (14)", `=IFERROR(VLOOKUP(${t}, CALCULATIONS!$A$3:$U, 19, 0), 0)`], 
    ["BOLLINGER %B", `=IFERROR(VLOOKUP(${t}, CALCULATIONS!$A$3:$U, 20, 0), 0)`] 
  ];

  chartSheet.getRange(5, 1, data.length, 1).setValues(data.map(r => [r[0]])).setFontWeight("bold");
  chartSheet.getRange(5, 2, data.length, 1).setFormulas(data.map(r => [r[1]]));
  
  [10, 18, 27].forEach(r => chartSheet.getRange(r, 1, 1, 2).setBackground("#444").setFontColor("white").setHorizontalAlignment("center"));
  
  SpreadsheetApp.flush();
  chartSheet.getRange("B5:B32").setHorizontalAlignment("left");
  chartSheet.getRangeList(["B6", "B7", "B11", "B13:B16", "B19:B22", "B25", "B28:B31"]).setNumberFormat("#,##0.00");
  chartSheet.getRangeList(["B8", "B12", "B32"]).setNumberFormat("0.00%");

  SpreadsheetApp.flush();
  updateDynamicChart();
}

/**
 * 4. CALC ENGINE
 */
function generateCalculationsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getSheetByName("DATA");
  const inputSheet = ss.getSheetByName("INPUT");
  if (!dataSheet) return;
  const tickers = getCleanTickers(inputSheet);
  let calcSheet = ss.getSheetByName("CALCULATIONS") || ss.insertSheet("CALCULATIONS");
  calcSheet.clear().clearFormats();

  calcSheet.getRange("A1:D1").merge().setValue("[ CORE IDENT ]").setBackground("#263238").setFontColor("white").setHorizontalAlignment("center").setFontWeight("bold");
  calcSheet.getRange("E1:G1").merge().setValue("[ PERFORMANCE ]").setBackground("#0D47A1").setFontColor("white").setHorizontalAlignment("center").setFontWeight("bold");
  calcSheet.getRange("H1:O1").merge().setValue("[ MOMENTUM ]").setBackground("#1B5E20").setFontColor("white").setHorizontalAlignment("center").setFontWeight("bold");
  calcSheet.getRange("P1:T1").merge().setValue("[ RISK LEVELS ]").setBackground("#B71C1C").setFontColor("white").setHorizontalAlignment("center").setFontWeight("bold");
  calcSheet.getRange("U1").setValue("[ ANALYST ]").setBackground("#424242").setFontColor("white").setHorizontalAlignment("center").setFontWeight("bold");

  const headers = [["Ticker", "Price", "Change %", "DECISION", "ATH (TRUE)", "ATH Diff %", "R:R Quality", "Trend Score", "Trend State", "SMA 20", "SMA 50", "SMA 200", "Vol Trend", "RSI", "Divergence", "Support", "Target (3:1)", "Resistance", "ATR (14)", "Bollinger %B", "REASONING"]];
  calcSheet.getRange(2, 1, 1, 21).setValues(headers).setBackground("#212121").setFontColor("white").setFontWeight("bold");
  
  const formulas = [];
  tickers.forEach((ticker, i) => {
    const rowNum = i + 3, tickerDataStart = (i * 7) + 1, closeCol = columnToLetter(tickerDataStart + 4), lastRow = `COUNTA(DATA!$${closeCol}:$${closeCol})`;
    formulas.push([
      `=ROUND(IFERROR(GOOGLEFINANCE("${ticker}", "price")), 2)`,
      `=IFERROR(GOOGLEFINANCE("${ticker}", "changepct")/100, 0)`,
      `=IFERROR(IFS(B${rowNum} < P${rowNum}, "STOP LOSS", OR(B${rowNum} >= Q${rowNum}, T${rowNum} > 1.0), "TAKE PROFIT", AND(T${rowNum}<0.15, B${rowNum}>P${rowNum}), "BUY DIP", AND(B${rowNum}>R${rowNum}, M${rowNum}>1.2), "BREAKOUT", AND(B${rowNum}>J${rowNum}, T${rowNum}<0.85), "RIDE TREND", TRUE, "HOLD"), "Wait")`,
      `=IFERROR(DATA!${columnToLetter(tickerDataStart + 1)}2, "-")`, 
      `=IFERROR((B${rowNum}-E${rowNum})/E${rowNum}, 0)`, 
      `=IFERROR(IF((Q${rowNum}-B${rowNum})/MAX(0.01, B${rowNum}-P${rowNum}) >= 3, "PRIME", "RISKY"), "—")`,
      `=REPT("★", (B${rowNum}>AVERAGE(OFFSET(DATA!$${closeCol}$3, ${lastRow}-20, 0, 20))) + (B${rowNum}>AVERAGE(OFFSET(DATA!$${closeCol}$3, ${lastRow}-50, 0, 50))))`,
      `=IF(B${rowNum}>AVERAGE(OFFSET(DATA!$${closeCol}$3, ${lastRow}-200, 0, 200)), "BULL REGIME", "BEAR REGIME")`,
      `=ROUND(IFERROR(AVERAGE(OFFSET(DATA!$${closeCol}$3, ${lastRow}-20, 0, 20)), 0), 2)`,
      `=ROUND(IFERROR(AVERAGE(OFFSET(DATA!$${closeCol}$3, ${lastRow}-50, 0, 50)), 0), 2)`,
      `=ROUND(IFERROR(AVERAGE(OFFSET(DATA!$${closeCol}$3, ${lastRow}-200, 0, 200)), 0), 2)`,
      `=ROUND(IFERROR(OFFSET(DATA!$${columnToLetter(tickerDataStart+5)}$3, ${lastRow}-1, 0) / AVERAGE(OFFSET(DATA!$${columnToLetter(tickerDataStart+5)}$3, ${lastRow}-21, 0, 20)), 1), 2)`,
      `=ROUND(IFERROR(100-(100/(1+(AVERAGEIF(ARRAYFORMULA(OFFSET(DATA!$${closeCol}$3, ${lastRow}-15, 0, 15)-OFFSET(DATA!$${closeCol}$3, ${lastRow}-16, 0, 15)),">0")/ABS(AVERAGEIF(ARRAYFORMULA(OFFSET(DATA!$${closeCol}$3, ${lastRow}-15, 0, 15)-OFFSET(DATA!$${closeCol}$3, ${lastRow}-16, 0, 15)),"<0"))))), 50), 2)`,
      // FIX: STABLE INDEX INSTEAD OF VOLATILE OFFSET FOR DIVERGENCE
      `=IFERROR(IFS(AND(B${rowNum} < INDEX(DATA!$${closeCol}:$${closeCol}, ${lastRow}-14), N${rowNum} > 50), "BULLISH DIV", AND(B${rowNum} > INDEX(DATA!$${closeCol}:$${closeCol}, ${lastRow}-14), N${rowNum} < 50), "BEARISH DIV", TRUE, "CONVERGENT"), "—")`,
      `=ROUND(IFERROR(MIN(OFFSET(DATA!$${columnToLetter(tickerDataStart+3)}$3, ${lastRow}-21, 0, 20)), 0), 2)`,
      `=ROUND(B${rowNum} + ((B${rowNum}-P${rowNum}) * 3), 2)`,
      `=ROUND(IFERROR(MAX(OFFSET(DATA!$${columnToLetter(tickerDataStart+2)}$3, ${lastRow}-51, 0, 50)), 0), 2)`,
      `=ROUND(IFERROR(AVERAGE(ARRAYFORMULA(OFFSET(DATA!$${columnToLetter(tickerDataStart+2)}$3, ${lastRow}-14, 0, 14)-OFFSET(DATA!$${columnToLetter(tickerDataStart+3)}$3, ${lastRow}-14, 0, 14))), 0), 2)`,
      `=ROUND(IFERROR(((B${rowNum}-AVERAGE(OFFSET(DATA!$${closeCol}$3, ${lastRow}-20, 0, 20))) / (4*STDEV(OFFSET(DATA!$${closeCol}$3, ${lastRow}-20, 0, 20)))) + 0.5, 0.5), 2)`,
      `=IFS(D${rowNum}="STOP LOSS", "🛑 STOP: Price $"&B${rowNum}&" broke Floor $"&P${rowNum}&". Trend: "&I${rowNum}&". RSI: "&N${rowNum}&".", D${rowNum}="TAKE PROFIT", "💰 PROFIT: Price $"&B${rowNum}&" hit Target $"&Q${rowNum}&" or Band Top (%B "&TEXT(T${rowNum}, "0.00")&"). ATR: "&S${rowNum}&".", D${rowNum}="BUY DIP", "🎯 DIP: Price $"&B${rowNum}&" > SMA200 ($"&L${rowNum}&"). Oversold (RSI "&N${rowNum}&", %B "&TEXT(T${rowNum}, "0.00")&"). Div: "&O${rowNum}&".", D${rowNum}="BREAKOUT", "🚀 BREAK: Price $"&B${rowNum}&" cleared Ceiling $"&R${rowNum}&". Vol: "&TEXT(M${rowNum}, "0.0")&"x. ATH: $"&E${rowNum}&". Trend: "&I${rowNum}&".", D${rowNum}="RIDE TREND", "🌊 RIDE: Holding SMA20 ($"&J${rowNum}&"). Div: "&O${rowNum}&". Next Res: $"&R${rowNum}&". Vol: "&TEXT(M${rowNum}, "0.0")&"x.", TRUE, "⏳ WAIT: Range $"&P${rowNum}&"-"&R${rowNum}&". Vol "&TEXT(M${rowNum}, "0.0")&"x. RSI "&N${rowNum}&". Trend: "&I${rowNum}&".")`
    ]);
  });
  calcSheet.getRange(3, 1, tickers.length, 1).setValues(tickers.map(t => [t]));
  calcSheet.getRange(3, 2, formulas.length, 20).setFormulas(formulas);
  
  SpreadsheetApp.flush();
  calcSheet.getRange("C3:C").setNumberFormat("0.00%");
  calcSheet.getRange("F3:F").setNumberFormat("0.00%");
  calcSheet.getRange("T3:T").setNumberFormat("0.00%");
}

function generateDataSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inputSheet = ss.getSheetByName("INPUT");
  if (!inputSheet) return;
  const tickers = getCleanTickers(inputSheet);
  let dataSheet = ss.getSheetByName("DATA") || ss.insertSheet("DATA");
  dataSheet.clear().clearFormats();
  
  dataSheet.getRange("G1").setValue("Last Update: " + Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "yyyy-MM-dd HH:mm")).setFontWeight("bold").setFontColor("blue");

  tickers.forEach((ticker, i) => {
    const colStart = (i * 7) + 1;
    dataSheet.getRange(1, colStart).setNumberFormat("@").setValue(ticker).setFontWeight("bold");
    dataSheet.getRange(2, colStart).setValue("ATH:");
    dataSheet.getRange(2, colStart + 1).setFormula(`=MAX(QUERY(GOOGLEFINANCE("${ticker}", "high", "1/1/2000", TODAY()), "SELECT Col2 LABEL Col2 ''"))`);
    dataSheet.getRange(3, colStart).setFormula(`=IFERROR(GOOGLEFINANCE("${ticker}", "all", TODAY()-800, TODAY()), "No Data")`);
    dataSheet.getRange(4, colStart, 1000, 1).setNumberFormat("yyyy-mm-dd");
    dataSheet.getRange(4, colStart + 1, 1000, 5).setNumberFormat("#,##0.00");
  });
}

function getCleanTickers(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return [];
  return sheet.getRange(3, 1, lastRow-2, 1).getValues().flat().filter(t => t && t.toString().trim() !== "").map(t => t.toString().toUpperCase());
}
function columnToLetter(column) {
  let temp, letter = '';
  while (column > 0) { temp = (column - 1) % 26; letter = String.fromCharCode(temp + 65) + letter; column = (column - temp - 1) / 26; }
  return letter;
}
function forceExpandSheet(sheet, targetCols) {
  if (sheet.getMaxColumns() < targetCols) sheet.insertColumnsAfter(sheet.getMaxColumns(), targetCols - sheet.getMaxColumns());
}