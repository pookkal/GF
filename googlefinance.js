/**
 * ==============================================================================
 * BASELINE LABEL: STABLE_MASTER_V33_MENU_UPDATE
 * DATE: 22 DEC 2025
 * FIX: Added "Build Calculations" to Menu. Preserved all V32 fixes.
 * ==============================================================================
 */

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📈 Institutional Terminal')
    .addItem('🚀 1-CLICK REBUILD ALL', 'FlushAllSheetsAndBuild')
    .addSeparator()
    .addItem('1. Fetch Data Only', 'generateDataSheet')
    .addItem('2. Build Calculations', 'generateCalculationsSheet') // NEW OPTION
    .addItem('3. Refresh Dashboard Only', 'generateDashboardSheet')
    .addItem('4. Setup Chart Only', 'setupChartSheet')
    .addToUi();
}

function FlushAllSheetsAndBuild() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetsToDelete = ["DATA", "CALCULATIONS", "CHART", "DASHBOARD"];
  const ui = SpreadsheetApp.getUi();
  if (ui.alert('🚨 Full Rebuild', 'Refresh entire terminal?', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;

  sheetsToDelete.forEach(name => {
    let sheet = ss.getSheetByName(name);
    if (sheet) ss.deleteSheet(sheet);
  });

  ui.showModelessDialog(HtmlService.createHtmlOutput("<b>1/4:</b> Syncing Global Data..."), "Status");
  generateDataSheet();
  SpreadsheetApp.flush();
  Utilities.sleep(2000); 

  ui.showModelessDialog(HtmlService.createHtmlOutput("<b>2/4:</b> Calculating Metrics..."), "Status");
  generateCalculationsSheet();
  SpreadsheetApp.flush();

  ui.showModelessDialog(HtmlService.createHtmlOutput("<b>3/4:</b> Updating Dashboard..."), "Status");
  generateDashboardSheet(); 
  SpreadsheetApp.flush();

  ui.showModelessDialog(HtmlService.createHtmlOutput("<b>4/4:</b> Drawing Chart..."), "Status");
  setupChartSheet();
  ui.alert('✅ Rebuild Complete', 'Terminal fully synchronized.', ui.ButtonSet.OK);
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
  const headers = [["Ticker", "Price", "Change %", "DECISION", "R:R Quality", "Trend Score", "Trend State", "SMA 20", "SMA 50", "SMA 200", "Vol Trend", "RSI", "Divergence", "Support", "Target (3:1)", "Resistance", "ATR (14)", "Bollinger %B", "REASONING"]];
  dashboard.getRange(2, 1, 1, 19).setValues(headers).setBackground("#212121").setFontColor("white").setFontWeight("bold");

  const formula = '=IFERROR(SORT(FILTER(CALCULATIONS!$A$3:$S, ISNUMBER(MATCH(CALCULATIONS!$A$3:$A, FILTER(INPUT!$A$3:$A, ' +
    '(IF(INPUT!$B$1="", 0, REGEXMATCH(INPUT!$B$3:$B, "(?i)\\b(" & SUBSTITUTE(SUBSTITUTE(INPUT!$B$1, ", ", "|"), ",", "|") & ")\\b"))) * ' +
    '(IF(OR(INPUT!$C$1="", INPUT!$C$1="ALL"), 1, REGEXMATCH(INPUT!$C$3:$C, "(?i)\\b(" & SUBSTITUTE(SUBSTITUTE(INPUT!$C$1, ", ", "|"), ",", "|") & ")\\b")))' +
    '), 0))), 3, FALSE), "Adjust B1/C1 Filters")';
  
  dashboard.getRange("A3").setFormula(formula);
  dashboard.getRange("A:S").setHorizontalAlignment("left");
  dashboard.getRange("C3:C200").setNumberFormat("0.00%");
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
  chartSheet.setColumnWidth(6, 500);

  chartSheet.getRange("A1").setValue("TICKER:").setFontWeight("bold");
  chartSheet.getRange("B1").setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(tickers).build()).setValue(tickers[0]).setBackground("#FFF9C4");
  chartSheet.getRange("D1").setValue("VIEW:").setFontWeight("bold");
  chartSheet.getRange("D2").setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(["DAILY", "WEEKLY"]).build()).setValue("DAILY").setBackground("#E1F5FE");
  
  chartSheet.getRange("E1").setValue("REASON:").setFontWeight("bold").setFontColor("#1565C0");
  chartSheet.getRange("F1").setFormula('=IFERROR(VLOOKUP(B1, CALCULATIONS!$A$3:$S, 19, 0), "—")').setWrap(true);

  chartSheet.getRange("A2:C2").setValues([["YEAR", "MONTH", "DAY"]]).setBackground("#222").setFontColor("#FFF").setHorizontalAlignment("center").setFontWeight("bold");
  const numRule = (max) => SpreadsheetApp.newDataValidation().requireValueInList(Array.from({length: max + 1}, (_, i) => i)).build();
  chartSheet.getRange("A3").setDataValidation(numRule(5)).setValue(0);
  chartSheet.getRange("B3").setDataValidation(numRule(12)).setValue(3);
  chartSheet.getRange("C3").setDataValidation(numRule(31)).setValue(0);
  chartSheet.getRange("A3:C3").setBackground("#F5F5F5").setHorizontalAlignment("center");

  chartSheet.getRange("A4").setValue("DATE").setFontWeight("bold").setBackground("#EEE");
  chartSheet.getRange("B4").setFormula("=DATE(YEAR(TODAY())-A3, MONTH(TODAY())-B3, DAY(TODAY())-C3)").setNumberFormat("yyyy-mm-dd").setFontColor("black");

  const t = "B1";
  
  // Data Mapping (Indices 0-27)
  const data = [
    ["SIGNAL (DECISION)", `=IFERROR(VLOOKUP(${t}, CALCULATIONS!$A$3:$S, 4, 0), "Wait")`], // Row 5
    ["LIVE PRICE", `=IFERROR(GOOGLEFINANCE(${t}, "price"), 0)`], // Row 6
    ["CHANGE ($)", `=IFERROR(B6 - GOOGLEFINANCE(${t}, "closeyest"), 0)`], // Row 7
    ["CHANGE (%)", `=IFERROR(GOOGLEFINANCE(${t}, "changepct")/100, 0)`], // Row 8
    ["", ""], // Row 9
    ["[ VALUATION METRICS ]", ""], // Row 10
    ["ALL TIME HIGH", `=IFERROR(VLOOKUP(${t}, CALCULATIONS!$A$3:$S, 16, 0), 0)`], 
    ["DIFF FROM ATH %", `=IFERROR((B6-B11)/B11, 0)`], 
    ["P/E RATIO", `=IFERROR(GOOGLEFINANCE(${t}, "pe"), 0)`],
    ["EPS", `=IFERROR(GOOGLEFINANCE(${t}, "eps"), 0)`],
    ["52W HIGH", `=IFERROR(GOOGLEFINANCE(${t}, "high52"), 0)`],
    ["52W LOW", `=IFERROR(GOOGLEFINANCE(${t}, "low52"), 0)`],
    ["", ""], // Row 17
    ["[ MOMENTUM & TREND ]", ""], // Row 18
    ["SMA 20", `=IFERROR(VLOOKUP(${t}, CALCULATIONS!$A$3:$S, 8, 0), 0)`],
    ["SMA 50", `=IFERROR(VLOOKUP(${t}, CALCULATIONS!$A$3:$S, 9, 0), 0)`],
    ["SMA 200", `=IFERROR(VLOOKUP(${t}, CALCULATIONS!$A$3:$S, 10, 0), 0)`],
    ["RSI (14)", `=IFERROR(VLOOKUP(${t}, CALCULATIONS!$A$3:$S, 12, 0), 50)`],
    ["TREND STATE", `=IFERROR(VLOOKUP(${t}, CALCULATIONS!$A$3:$S, 7, 0), "—")`],
    ["DIVERGENCE", `=IFERROR(VLOOKUP(${t}, CALCULATIONS!$A$3:$S, 13, 0), "Neutral")`], // Row 24
    ["RELATIVE VOLUME", `=IFERROR(VLOOKUP(${t}, CALCULATIONS!$A$3:$S, 11, 0), 1)`],
    ["", ""], // Row 26
    ["[ TECHNICAL LEVELS ]", ""], // Row 27
    ["SUPPORT FLOOR", `=IFERROR(VLOOKUP(${t}, CALCULATIONS!$A$3:$S, 14, 0), 0)`], // Row 28
    ["RESISTANCE CEILING", `=IFERROR(VLOOKUP(${t}, CALCULATIONS!$A$3:$S, 16, 0), 0)`], // Row 29
    ["TARGET (3:1 R:R)", `=IFERROR(VLOOKUP(${t}, CALCULATIONS!$A$3:$S, 15, 0), 0)`], 
    ["ATR (14)", `=IFERROR(VLOOKUP(${t}, CALCULATIONS!$A$3:$S, 17, 0), 0)`], 
    ["BOLLINGER %B", `=IFERROR(VLOOKUP(${t}, CALCULATIONS!$A$3:$S, 18, 0), 0)`] // Row 32
  ];

  chartSheet.getRange(5, 1, data.length, 1).setValues(data.map(r => [r[0]])).setFontWeight("bold");
  chartSheet.getRange(5, 2, data.length, 1).setFormulas(data.map(r => [r[1]]));
  
  [10, 18, 27].forEach(r => chartSheet.getRange(r, 1, 1, 2).setBackground("#444").setFontColor("white").setHorizontalAlignment("center"));
  
  // Formatting
  chartSheet.getRange("B8, B12, B32").setNumberFormat("0.00%");
  chartSheet.getRange("B5:B32").setHorizontalAlignment("left");
  chartSheet.getRange("B6, B7, B11, B13:B16, B19:B22, B25, B28:B31").setNumberFormat("#,##0.00");

  SpreadsheetApp.flush();
  updateDynamicChart();
}

/**
 * 3. CHART ENGINE (CHART TRIGGER FIXED)
 */
function updateDynamicChart() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("CHART");
  const dataSheet = ss.getSheetByName("DATA");
  if (!sheet || !dataSheet) return;

  const ticker = sheet.getRange("B1").getValue();
  const startDate = sheet.getRange("B4").getValue();
  const isWeekly = sheet.getRange("D2").getValue() === "WEEKLY";

  SpreadsheetApp.flush(); // FORCE SYNC
  
  const supportVal = Number(sheet.getRange("B28").getValue()) || 0; // Row 28
  const resistanceVal = Number(sheet.getRange("B29").getValue()) || 0; // Row 29

  const rawHeaders = dataSheet.getRange(1, 1, 1, dataSheet.getLastColumn()).getValues()[0];
  const colIdx = rawHeaders.indexOf(ticker);
  if (colIdx === -1) return;

  const rawData = dataSheet.getRange(1, colIdx + 1, dataSheet.getLastRow(), 6).getValues();
  let masterData = [], viewVols = [], prices = [];

  for (let i = 2; i < rawData.length; i++) {
    let row = rawData[i], d = row[0], close = Number(row[4]), vol = Number(row[5]);
    if (!d || !(d instanceof Date) || isNaN(close) || close < 0.01) continue;
    if (d < startDate) continue;
    if (isWeekly && d.getDay() !== 5) continue;

    let slice = rawData.slice(Math.max(2, i-200), i+1).map(r => r[4]);
    let s20 = slice.length >= 20 ? Number((slice.slice(-20).reduce((a,b)=>a+b,0)/20).toFixed(2)) : null;
    let s50 = slice.length >= 50 ? Number((slice.slice(-50).reduce((a,b)=>a+b,0)/50).toFixed(2)) : null;
    let s200 = slice.length >= 200 ? Number((slice.slice(-200).reduce((a,b)=>a+b,0)/200).toFixed(2)) : null;

    masterData.push([Utilities.formatDate(d, ss.getSpreadsheetTimeZone(), "MMM dd"), close, (close >= (i>2?rawData[i-1][4]:close))?vol:null, (close < (i>2?rawData[i-1][4]:close))?vol:null, s20, s50, s200, resistanceVal, supportVal]);
    viewVols.push(vol); prices.push(close);
    if(s20) prices.push(s20); if(s50) prices.push(s50); if(s200) prices.push(s200);
  }

  if (masterData.length === 0) return;
  if (supportVal > 0) prices.push(supportVal); if (resistanceVal > 0) prices.push(resistanceVal);

  const minP = Math.min(...prices.filter(p => p > 0)) * 0.98;
  const maxP = Math.max(...prices.filter(p => p > 0)) * 1.02;
  const maxVol = Math.max(...viewVols);

  const chartLabels = [["Date", "Price", "Bull Vol", "Bear Vol", "SMA 20", "SMA 50", "SMA 200", "Resistance", "Support"]];
  sheet.getRange(2, 26, 1, 9).setValues(chartLabels).setFontWeight("bold").setFontColor("white");
  sheet.getRange(3, 26, 1500, 9).clearContent();
  sheet.getRange(3, 26, masterData.length, 9).setValues(masterData);
  SpreadsheetApp.flush();

  sheet.getCharts().forEach(c => sheet.removeChart(c));
  const chart = sheet.newChart()
    .setChartType(Charts.ChartType.COMBO)
    .addRange(sheet.getRange(2, 26, masterData.length + 1, 9))
    .setOption('useFirstRowAsHeaders', true)
    .setOption('series', {
      0: {type: 'line', color: '#1A73E8', lineWidth: 3, labelInLegend: 'Price'},
      1: {type: 'bars', color: '#2E7D32', targetAxisIndex: 1, labelInLegend: 'Bull Vol'},
      2: {type: 'bars', color: '#C62828', targetAxisIndex: 1, labelInLegend: 'Bear Vol'},
      3: {type: 'line', color: '#FBBC04', lineWidth: 1.5, labelInLegend: 'SMA 20'},
      4: {type: 'line', color: '#9C27B0', lineWidth: 1.5, labelInLegend: 'SMA 50'},
      5: {type: 'line', color: '#FF9800', lineWidth: 2, labelInLegend: 'SMA 200'},
      6: {type: 'line', color: '#B71C1C', lineDashStyle: [4, 4], labelInLegend: 'Resistance'},
      7: {type: 'line', color: '#0D47A1', lineDashStyle: [4, 4], labelInLegend: 'Support'}
    })
    .setOption('vAxes', {
      0: {viewWindow: {min: minP, max: maxP}, gridlines: {color: '#f3f3f3'}}, 
      1: {viewWindow: {min: 0, max: maxVol * 8}, textStyle: {color: 'none'}, gridlines: {count: 0}}
    })
    .setOption('legend', {position: 'top', textStyle: {fontSize: 10}})
    .setPosition(4, 3, 10, 10).setOption('width', 1150).setOption('height', 650).build();
  sheet.insertChart(chart);
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
  calcSheet.getRange("H1:M1").merge().setValue("[ MOMENTUM & TREND ]").setBackground("#1B5E20").setFontColor("white").setHorizontalAlignment("center").setFontWeight("bold");
  calcSheet.getRange("N1:R1").merge().setValue("[ VOLATILITY LEVELS ]").setBackground("#B71C1C").setFontColor("white").setHorizontalAlignment("center").setFontWeight("bold");
  calcSheet.getRange("S1").setValue("[ ANALYST ]").setBackground("#424242").setFontColor("white").setHorizontalAlignment("center").setFontWeight("bold");

  const headers = [["Ticker", "Price", "Change %", "DECISION", "R:R Quality", "Trend Score", "Trend State", "SMA 20", "SMA 50", "SMA 200", "Vol Trend", "RSI", "Divergence", "Support", "Target (3:1)", "Resistance", "ATR (14)", "Bollinger %B", "REASONING"]];
  calcSheet.getRange(2, 1, 1, 19).setValues(headers).setBackground("#212121").setFontColor("white").setFontWeight("bold");
  
  const formulas = [];
  tickers.forEach((ticker, i) => {
    const rowNum = i + 3, tickerDataStart = (i * 7) + 1, closeCol = columnToLetter(tickerDataStart + 4), lastRow = `COUNTA(DATA!$${closeCol}:$${closeCol})`;
    formulas.push([
      `=ROUND(IFERROR(GOOGLEFINANCE("${ticker}", "price")), 2)`,
      `=IFERROR(GOOGLEFINANCE("${ticker}", "changepct")/100, 0)`,
      `=IFERROR(IFS(B${rowNum} < N${rowNum}, "🚨 EXIT", B${rowNum} >= O${rowNum}, "💰 PROFIT", AND(R${rowNum}<0.2, L${rowNum}<45), "🎯 BUY DIP", AND(B${rowNum}>I${rowNum}, K${rowNum}>1.2), "🚀 STRONG BUY", TRUE, "⚖️ HOLD"), "Wait")`,
      `=IFERROR(IF((O${rowNum}-B${rowNum})/MAX(0.01, B${rowNum}-N${rowNum}) >= 3, "💎 HIGH", "⚖️ MED"), "—")`,
      `=REPT("★", (B${rowNum}>AVERAGE(OFFSET(DATA!$${closeCol}$1, ${lastRow}-20, 0, 20))) + (B${rowNum}>AVERAGE(OFFSET(DATA!$${closeCol}$1, ${lastRow}-50, 0, 50))))`,
      `=IF(B${rowNum}>AVERAGE(OFFSET(DATA!$${closeCol}$1, ${lastRow}-200, 0, 200)), "🚀 BULLISH", "📉 BEARISH")`,
      `=ROUND(IFERROR(AVERAGE(OFFSET(DATA!$${closeCol}$1, ${lastRow}-20, 0, 20)), 0), 2)`,
      `=ROUND(IFERROR(AVERAGE(OFFSET(DATA!$${closeCol}$1, ${lastRow}-50, 0, 50)), 0), 2)`,
      `=ROUND(IFERROR(AVERAGE(OFFSET(DATA!$${closeCol}$1, ${lastRow}-200, 0, 200)), 0), 2)`,
      `=ROUND(IFERROR(OFFSET(DATA!$${columnToLetter(tickerDataStart+5)}$1, ${lastRow}-1, 0) / AVERAGE(OFFSET(DATA!$${columnToLetter(tickerDataStart+5)}$1, ${lastRow}-21, 0, 20)), 1), 2)`,
      `=ROUND(IFERROR(100-(100/(1+(AVERAGEIF(ARRAYFORMULA(OFFSET(DATA!$${closeCol}$1, ${lastRow}-15, 0, 15)-OFFSET(DATA!$${closeCol}$1, ${lastRow}-16, 0, 15)),">0")/ABS(AVERAGEIF(ARRAYFORMULA(OFFSET(DATA!$${closeCol}$1, ${lastRow}-15, 0, 15)-OFFSET(DATA!$${closeCol}$1, ${lastRow}-16, 0, 15)),"<0"))))), 50), 2)`,
      `=IFERROR(IFS(AND(B${rowNum}>OFFSET(DATA!$${closeCol}$1,${lastRow}-11,0), L${rowNum}<OFFSET(DATA!$${closeCol}$1,${lastRow}-11,0,1,2)), "📉 BEARISH", AND(B${rowNum}<OFFSET(DATA!$${closeCol}$1,${lastRow}-11,0), L${rowNum}>OFFSET(DATA!$${closeCol}$1,${lastRow}-11,0,1,2)), "🚀 BULLISH", TRUE, "Neutral"), "—")`,
      `=ROUND(IFERROR(MIN(OFFSET(DATA!$${columnToLetter(tickerDataStart+3)}$1, ${lastRow}-21, 0, 20)), 0), 2)`,
      `=ROUND(B${rowNum} + ((B${rowNum}-N${rowNum}) * 3), 2)`,
      `=ROUND(IFERROR(MAX(OFFSET(DATA!$${columnToLetter(tickerDataStart+2)}$1, ${lastRow}-51, 0, 50)), 0), 2)`,
      `=ROUND(IFERROR(AVERAGE(ARRAYFORMULA(OFFSET(DATA!$${columnToLetter(tickerDataStart+2)}$1, ${lastRow}-14, 0, 14)-OFFSET(DATA!$${columnToLetter(tickerDataStart+3)}$1, ${lastRow}-14, 0, 14))), 0), 2)`,
      `=ROUND(IFERROR((B${rowNum}-AVERAGE(OFFSET(DATA!$${closeCol}$1, ${lastRow}-20, 0, 20))) / (4*STDEV(OFFSET(DATA!$${closeCol}$1, ${lastRow}-20, 0, 20))), 0.5), 2)`,
      `="Trend Regime for "&A${rowNum}&" is "&G${rowNum}&". Institutional floor sits at $"&N${rowNum}&"."`
    ]);
  });
  calcSheet.getRange(3, 1, tickers.length, 1).setValues(tickers.map(t => [t]));
  calcSheet.getRange(3, 2, formulas.length, 18).setFormulas(formulas);
}

function generateDataSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inputSheet = ss.getSheetByName("INPUT");
  if (!inputSheet) return;
  const tickers = getCleanTickers(inputSheet);
  let dataSheet = ss.getSheetByName("DATA") || ss.insertSheet("DATA");
  dataSheet.clear().clearFormats();
  tickers.forEach((ticker, i) => {
    const colStart = (i * 7) + 1;
    dataSheet.getRange(1, colStart).setNumberFormat("@").setValue(ticker).setFontWeight("bold");
    dataSheet.getRange(2, colStart).setFormula(`=IFERROR(GOOGLEFINANCE("${ticker}", "all", TODAY()-800, TODAY()), "No Data")`);
    dataSheet.getRange(3, colStart, 1000, 1).setNumberFormat("yyyy-mm-dd");
    dataSheet.getRange(3, colStart + 1, 1000, 5).setNumberFormat("#,##0.00");
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